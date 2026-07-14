import io
import re
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
    """Create a formatted Excel workbook entirely in memory."""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Web Attributes")

        workbook = writer.book
        worksheet = writer.sheets["Web Attributes"]

        header_format = workbook.add_format(
            {
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": "#008577",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        )
        text_format = workbook.add_format({"num_format": "@"})
        true_format = workbook.add_format(
            {"bg_color": "#E2F0D9", "font_color": "#375623"}
        )
        false_format = workbook.add_format(
            {"bg_color": "#FCE4D6", "font_color": "#843C0C"}
        )

        for column_number, column_name in enumerate(df.columns):
            worksheet.write(0, column_number, column_name, header_format)

        # Keep SKU as text so values such as 062994 retain their leading zero.
        worksheet.set_column("A:A", 16, text_format)
        worksheet.set_column("B:B", 66)
        worksheet.set_column("C:C", 24)
        worksheet.set_column("D:D", 12)
        worksheet.set_column("E:E", 10)
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

        if not df.empty:
            worksheet.conditional_format(
                1,
                3,
                len(df),
                3,
                {
                    "type": "text",
                    "criteria": "containing",
                    "value": "TRUE",
                    "format": true_format,
                },
            )
            worksheet.conditional_format(
                1,
                3,
                len(df),
                3,
                {
                    "type": "text",
                    "criteria": "containing",
                    "value": "FALSE",
                    "format": false_format,
                },
            )

        worksheet.add_table(
            0,
            0,
            len(df),
            len(df.columns) - 1,
            {
                "name": "WebAttributeUpload",
                "columns": [{"header": column} for column in df.columns],
                "style": "Table Style Medium 4",
            },
        )

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
