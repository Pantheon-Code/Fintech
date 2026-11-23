# validation.py
import openpyxl

def load_valid_ids(sheet, id_col=1):
    """
    Reads a single column from an openpyxl worksheet and returns a set of valid IDs.
    Assumes row 1 is a header.
    """
    valid_ids = set()
    for row in sheet.iter_rows(min_row=2, values_only=True):
        account_id = row[id_col - 1]
        if account_id is not None:
            # Convert to string and strip whitespace to prevent "1010 " vs "1010" bugs
            valid_ids.add(str(account_id).strip())
    return valid_ids


def validate_coa_mappings(source_sheet, target_valid_ids, source_col=1, target_col=3):
    """
    Checks a COA sheet (e.g., Company A) to ensure:
    1. There are no duplicate source account IDs (ambiguous mappings).
    2. Every source account is mapped to a target ID (unmapped error).
    3. Every mapped target ID actually exists in the target COA (invalid ID error).
    
    Returns a dictionary of valid mappings {source_id: target_id} and a list of error strings.
    """
    mappings = {}
    errors = []

    for row_idx, row in enumerate(source_sheet.iter_rows(min_row=2, values_only=True), start=2):
        raw_source = row[source_col - 1]
        raw_target = row[target_col - 1]

        # Skip completely empty rows
        if raw_source is None and raw_target is None:
            continue

        source_id = str(raw_source).strip() if raw_source is not None else ""
        target_id = str(raw_target).strip() if raw_target is not None else ""

        # Check 1: Missing source ID
        if not source_id:
            errors.append(f"Row {row_idx} in '{source_sheet.title}': Missing source Account ID.")
            continue

        # Check 2: Ambiguous (duplicate) mapping
        if source_id in mappings:
            errors.append(f"Row {row_idx} in '{source_sheet.title}': Duplicate/ambiguous mapping for Account '{source_id}'.")
            continue

        # Check 3: Unmapped account
        if not target_id:
            errors.append(f"Row {row_idx} in '{source_sheet.title}': Account '{source_id}' is unmapped.")
            continue

        # Check 4: Target ID does not exist in the Universal COA
        if target_id not in target_valid_ids:
            errors.append(f"Row {row_idx} in '{source_sheet.title}': Mapped ID '{target_id}' for Account '{source_id}' does not exist in target COA.")
            continue

        # If all checks pass, store the clean mapping
        mappings[source_id] = target_id

    return mappings, errors