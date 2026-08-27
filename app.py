import re
import pandas as pd


def extract_search_key(text):
    """Extracts from volume (e.g. 100ML, 50ML, 20ML) up to the product keywords.

    Example: 'FG-PURPLLE-PER-100ML-AURA-LOVE STRUCK DELIGHT X 36' -> '100ML AURA
    LOVE STRUCK DELIGHT'
    """
    if not isinstance(text, str):
        return None

    # Find where the volume starts (e.g., 100ML, 50 ML, 20ML) and get everything after it
    match = re.search(r"(\d+\s*ML.*)", text, re.IGNORECASE)
    if not match:
        return None

    extracted = match.group(1)

    # Remove non-alphanumeric trailing tokens like 'X 36', 'X36', '-X-36' at the end
    extracted = re.sub(
        r"[\s\-_]+X[\s\-_]*\d+.*$", "", extracted, flags=re.IGNORECASE
    )

    # Normalize hyphens and multiple spaces into clean single spaces
    clean_key = re.sub(r"[\-_]+", " ", extracted).strip().upper()

    return clean_key


def map_ean_and_sku(df_invoices, df_master):
    """df_invoices: DataFrame containing 'Description of Goods' df_master:

    DataFrame containing master columns ('SKU Name', 'EAN', 'SKU Code')
    """
    # Create normalized search keys for matching
    df_invoices["Search_Key"] = df_invoices["Description of Goods"].apply(
        extract_search_key
    )

    # Prepare master list keys (extract similar patterns from master SKU names if needed, or normalize)
    df_master["Master_Search_Key"] = df_master["SKU Name"].apply(
        extract_search_key
    )

    # Merge based on extracted key
    merged_df = pd.merge(
        df_invoices,
        df_master[["Master_Search_Key", "EAN", "SKU Code", "SKU Name"]],
        left_on="Search_Key",
        right_on="Master_Search_Key",
        how="left",
        suffixes=("", "_Master"),
    )

    # Drop helper merge columns
    merged_df.drop(columns=["Search_Key", "Master_Search_Key"], inplace=True)

    return merged_df


# --- Example Usage ---
# df_invoices = pd.read_excel('invoices.xlsx')
# df_master = pd.read_excel('EAN_Master.xlsx')
# result_df = map_ean_and_sku(df_invoices, df_master)
