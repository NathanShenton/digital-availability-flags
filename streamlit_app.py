import io
import re
import zipfile
from xml.sax.saxutils import escape
from datetime import datetime

import pandas as pd
import streamlit as st


# -----------------------------------------------------------------------------
# Page setup
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Web Attribute Bulk Builder",
    page_icon="🛒",
    layout="wide",
)

# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------
PREFIX = "27-FINISHED GOODS"
COUNTRIES = {
    "GB": "United Kingdom",
    "IE": "Ireland",
}

ATTRIBUTE_GROUPS = {
    "Visibility": [
        ("is_active_for_web", "Active"),
        ("searchable_for_web", "Searchable"),
        ("coming_soon_for_web", "Coming soon"),
    ],
    "Purchase options": [
        ("purchasable_for_web", "Purchasable"),
        ("sold_online_for_web", "Sold online"),
        ("sold_in_store_for_web", "Sold in store"),
    ],
    "Subscription": [
        ("subscription_eligibility_for_web", "Eligible for subscription"),
        (
            "subscription_discontinued_for_web",
            "Discontinued for subscription",
        ),
    ],
    "Notifications": [
        ("email_me_when_available_for_web", "Email me when available"),
    ],
}

STATE_OPTIONS = ["No change", "TRUE", "FALSE"]


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def initialise_state() -> None:
    """Create a stable session-state value for every country/attribute pair."""
    for country_code in COUNTRIES:
        for attributes in ATTRIBUTE_GROUPS.values():
            for attribute_key, _ in attributes:
                state_key = f"state_{country_code}_{attribute_key}"
                if state_key not in st.session_state:
                    st.session_state[state_key] = "No change"


def parse_skus(raw_text: str) -> tuple[list[str], list[str]]:
    """
    Parse SKUs from lines, commas, semicolons or tabs.

    SKUs are deliberately kept as strings so leading zeroes are preserved.
    Blank values are ignored and duplicates are removed while retaining the
    original order.
    """
    tokens = re.split(r"[\n,;\t]+", raw_text)

    valid_skus: list[str] = []
    rejected_values: list[str] = []
    seen: set[str] = set()

    for token in tokens:
        sku = token.strip()
        if not sku:
            continue

        # Allow alphanumeric SKUs plus common safe separators. This prevents
        # pasted headings or narrative text entering the upload by accident.
        if not re.fullmatch(r"[A-Za-z0-9._-]+", sku):
            rejected_values.append(sku)
            continue

        if sku not in seen:
            valid_skus.append(sku)
            seen.add(sku)

    return valid_skus, rejected_values


def set_all_for_country(country_code: str, value: str) -> None:
    """Set all attributes for one country to the selected tri-state value."""
    for attributes in ATTRIBUTE_GROUPS.values():
        for attribute_key, _ in attributes:
            st.session_state[f"state_{country_code}_{attribute_key}"] = value


def build_output(skus: list[str]) -> pd.DataFrame:
    """Create one row per SKU, selected attribute and selected country."""
    rows: list[dict[str, str]] = []

    for country_code in COUNTRIES:
        for attributes in ATTRIBUTE_GROUPS.values():
            for attribute_key, display_name in attributes:
                selected_value = st.session_state[
                    f"state_{country_code}_{attribute_key}"
                ]

                if selected_value == "No change":
                    continue

                attribute_label = f"{attribute_key} ({display_name})"
                for sku in skus:
                    rows.append(
                        {
                            "sku": sku,
                            "attribute": attribute_label,
                            "prefix": PREFIX,
                            "value": selected_value,
                            "country": country_code,
                        }
                    )

    return pd.DataFrame(
        rows,
        columns=["sku", "attribute", "prefix", "value", "country"],
    )


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Create a valid .xlsx workbook without optional Excel packages."""
    output = io.BytesIO()

    def excel_column_name(column_number: int) -> str:
        result = ""
        while column_number:
            column_number, remainder = divmod(column_number - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def inline_string_cell(reference: str, value: object, style_id: int = 0) -> str:
        text = escape("" if value is None else str(value))
        style = f' s="{style_id}"' if style_id else ""
        return (
            f'<c r="{reference}" t="inlineStr"{style}>'
            f'<is><t xml:space="preserve">{text}</t></is></c>'
        )

    columns = list(df.columns)
    row_count = len(df) + 1
    last_column = excel_column_name(len(columns))
    worksheet_rows: list[str] = []

    header_cells = [
        inline_string_cell(f"{excel_column_name(index)}1", column, 1)
        for index, column in enumerate(columns, start=1)
    ]
    worksheet_rows.append(
        f'<row r="1" ht="22" customHeight="1">{"".join(header_cells)}</row>'
    )

    for excel_row, values in enumerate(df.itertuples(index=False, name=None), start=2):
        cells: list[str] = []
        for column_number, value in enumerate(values, start=1):
            style_id = 0
            if column_number == 1:
                style_id = 2
            elif column_number == 4:
                style_id = 3 if str(value) == "TRUE" else 4

            reference = f"{excel_column_name(column_number)}{excel_row}"
            cells.append(inline_string_cell(reference, value, style_id))

        worksheet_rows.append(f'<row r="{excel_row}">{"".join(cells)}</row>')

    worksheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
        '</sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols>'
        '<col min="1" max="1" width="16" customWidth="1"/>'
        '<col min="2" max="2" width="66" customWidth="1"/>'
        '<col min="3" max="3" width="24" customWidth="1"/>'
        '<col min="4" max="4" width="12" customWidth="1"/>'
        '<col min="5" max="5" width="10" customWidth="1"/>'
        '</cols>'
        f'<sheetData>{"".join(worksheet_rows)}</sheetData>'
        f'<autoFilter ref="A1:{last_column}{row_count}"/>'
        '</worksheet>'
    )

    content_types_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

    root_relationships_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

    workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Web Attributes" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''

    workbook_relationships_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="@"/></numFmts>
  <fonts count="4">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font>
    <font><color rgb="FF375623"/><sz val="11"/><name val="Calibri"/></font>
    <font><color rgb="FF843C0C"/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="5">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF008577"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    created = datetime.utcnow().replace(microsecond=0).isoformat() + "Z"
    core_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Streamlit</dc:creator>
  <cp:lastModifiedBy>Streamlit</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>
</cp:coreProperties>'''

    app_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Streamlit</Application>
</Properties>'''

    with zipfile.ZipFile(output, mode="w", compression=zipfile.ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", content_types_xml)
        workbook.writestr("_rels/.rels", root_relationships_xml)
        workbook.writestr("xl/workbook.xml", workbook_xml)
        workbook.writestr("xl/_rels/workbook.xml.rels", workbook_relationships_xml)
        workbook.writestr("xl/worksheets/sheet1.xml", worksheet_xml)
        workbook.writestr("xl/styles.xml", styles_xml)
        workbook.writestr("docProps/core.xml", core_xml)
        workbook.writestr("docProps/app.xml", app_xml)

    output.seek(0)
    return output.getvalue()

def render_country_controls(country_code: str, country_name: str) -> None:
    """Render the controls for one country in a layout similar to the PIM UI."""
    st.markdown(f"### {country_code} · {country_name}")

    action_col1, action_col2, action_col3, spacer = st.columns(
        [1.15, 1.15, 1.15, 5]
    )
    with action_col1:
        if st.button(
            "Set all TRUE",
            key=f"all_true_{country_code}",
            use_container_width=True,
        ):
            set_all_for_country(country_code, "TRUE")
            st.rerun()
    with action_col2:
        if st.button(
            "Set all FALSE",
            key=f"all_false_{country_code}",
            use_container_width=True,
        ):
            set_all_for_country(country_code, "FALSE")
            st.rerun()
    with action_col3:
        if st.button(
            "Clear country",
            key=f"clear_{country_code}",
            use_container_width=True,
        ):
            set_all_for_country(country_code, "No change")
            st.rerun()

    group_columns = st.columns(4)

    for column, (group_name, attributes) in zip(
        group_columns, ATTRIBUTE_GROUPS.items()
    ):
        with column:
            with st.container(border=True):
                st.markdown(f"**{group_name.upper()}**")

                for attribute_key, display_name in attributes:
                    st.selectbox(
                        display_name,
                        options=STATE_OPTIONS,
                        key=f"state_{country_code}_{attribute_key}",
                        help=(
                            "No change excludes this attribute from the output. "
                            "TRUE or FALSE creates an upload row for every SKU."
                        ),
                    )

    st.divider()


# -----------------------------------------------------------------------------
# App
# -----------------------------------------------------------------------------
initialise_state()

st.title("Web Attribute Bulk Builder")
st.caption(
    "Paste SKU IDs, choose TRUE or FALSE for the required GB and IE web "
    "attributes, then download the bulk-upload Excel file."
)

with st.container(border=True):
    st.subheader("1. Paste SKU IDs")
    raw_skus = st.text_area(
        "SKU IDs",
        height=180,
        placeholder="062994\n11111\n22222",
        help="Paste one SKU per line, or separate them with commas, semicolons or tabs.",
        label_visibility="collapsed",
    )

    skus, rejected = parse_skus(raw_skus)
    metric_col1, metric_col2, metric_col3 = st.columns(3)
    metric_col1.metric("Valid unique SKUs", len(skus))
    metric_col2.metric("Rejected values", len(rejected))
    metric_col3.metric("Upload prefix", PREFIX)

    if rejected:
        st.warning(
            "These values were ignored because they contain unsupported "
            f"characters: {', '.join(rejected[:10])}"
            + (" …" if len(rejected) > 10 else "")
        )

st.subheader("2. Select attribute values")
st.info(
    "Leave an attribute as **No change** to omit it from the output. "
    "Each chosen value applies to every pasted SKU for that country."
)

for code, name in COUNTRIES.items():
    render_country_controls(code, name)

st.subheader("3. Review and download")
output_df = build_output(skus)

summary_col1, summary_col2, summary_col3 = st.columns(3)
summary_col1.metric("Output rows", len(output_df))
summary_col2.metric(
    "Selected country/attribute combinations",
    0 if not skus else len(output_df) // len(skus),
)
summary_col3.metric("Countries included", output_df["country"].nunique() if not output_df.empty else 0)

if not skus:
    st.warning("Paste at least one valid SKU to create the file.")
elif output_df.empty:
    st.warning("Choose TRUE or FALSE for at least one attribute.")
else:
    st.dataframe(
        output_df,
        use_container_width=True,
        hide_index=True,
        height=420,
    )

    excel_bytes = dataframe_to_excel_bytes(output_df)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")

    st.download_button(
        label="Download Excel upload file",
        data=excel_bytes,
        file_name=f"web_attribute_bulk_upload_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
