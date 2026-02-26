from openai import OpenAI
import pandas as pd
import os

def main(fileName, receiveDate, vendorName):
    try:
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise ValueError("OPENAI_API_KEY environment variable is not set")
        client = OpenAI(api_key=api_key)

        file = client.files.create(
            file=open(f"packing_slip/{fileName}", "rb"),
            purpose="user_data"
        )

        prompt_text = f"""
Extract the shipment date from a PDF and output ONLY a single value (YYYY-MM-DD or "Not found") with no explanation or extra text. The PDF contains shipment information and a Receive Date: {receiveDate}. Vendor: {vendorName}. Use OCR and document-understanding to locate the Ship Date, attempting recognition even if text is blurred, noisy, scanned at low quality, tilted, rotated, partially occluded, or on multiple pages; allow minor OCR errors (e.g., "JAN" misread as "JAH" or "J4N", digits misrecognized, slashes/hyphens swapped, extra spaces) and automatically correct them only if the intended date is unambiguous; do NOT invent, estimate, or interpolate missing parts. If Westburne, the Ship Date is guaranteed to exist; find "Ship Date" (case-insensitive) and extract the date directly below it, correcting minor OCR errors if necessary, in format YY MMM DD (e.g., "25 JAN 12" or "2025/01/12"); search all pages if needed. If WESCO, packing slip is landscape; after rotation, find date at top-left corner in format MM-DD-YY or YY MMM DD. If year is missing, assume 2025. Convert found date to YYYY-MM-DD: YY MMM DD → 20YY-MM-DD, MM-DD-YY → 20YY-MM-DD. Ensure date ≤ Receive Date and ≥ 2025-04-01; if not, output "Not found". Output strictly one value only: YYYY-MM-DD or "Not found". Examples: '25 JAN 12' → 2025-01-12, '03-24-25' → 2025-03-24, missing/unreadable date → Not found. Do not include spaces, notes, or reasoning.
    """
        response = client.responses.create(
            model="gpt-5.2",
            input=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_file",
                            "file_id": file.id,
                        },
                        {
                            "type": "input_text",
                            "text": prompt_text,
                        },
                    ]
                }
            ]
        )
        print(f"{fileName} - {response.output_text}")
        return(response.output_text)
    except Exception as e:
        print(f"Error in checking file {fileName}")

def temp():

    df = pd.read_excel("packing_slip/output.xlsx", sheet_name="Sheet")

    for idx, row in df.iterrows():
        vendor = str(row.get("Vendor"))
        diff = row.get("Days Difference")
        if idx > 2000:
            df.to_excel("output.xlsx", sheet_name="Sheet", index=False)            
            break
        if vendor.startswith("Westburne") and (diff < 0 or diff > 10):
            packing_list = row.get("Packing List")
            receiving_date = row.get("Receiving Date")
            if pd.notna(packing_list) and pd.notna(receiving_date):
                shipping_date = main(packing_list, receiving_date)
                df.at[idx, "Shipping Date"] = shipping_date

    df.to_excel("output.xlsx", sheet_name="Sheet", index=False)

if __name__ == "__main__":
    temp()


'''
        prompt_text = f"""
        Extract the shipment date from a PDF following the precise instructions below, and output ONLY a single value (YYYY-MM-DD or "Not found") with no explanation, formatting, or extra text. ## Task Objective and Steps 1. **Input:** You will receive a PDF document containing shipment information and a "Receive Date" in YYYY-MM-DD format (e.g., 2025-06-12). 2. **Company Identification:** - If "Westburne": Find "Ship Date" (case-insensitive). The ship date is directly below this label, format is `YY MMM DD` (e.g., "25 JAN 12" or "2025/01/12"). - If "WESCO": Packing slip is landscape and must be rotated upright. Find the date at the top-left corner, format is `MM-DD-YY` or `YY MMM DD`. - Dates on the same page are in the same format. 3. **Date Extraction:** - Locate the date directly below "Ship Date" (case-insensitive) for Westburne, or the top-left for WESCO. - Only use explicitly present, human-readable dates; **never invent, estimate, or interpolate dates**. - If a date does not include a year, assume year equals 2025. 4. **Date Conversion:** - Convert the found date to YYYY-MM-DD: - For `YY MMM DD`: YY → 20YY, map MMM to MM, keep DD. Example: "25 JAN 12" → "2025-01-12". - For `MM-DD-YY`: MM → month, DD → day, YY → year (20YY). - For ambiguous forms (e.g., 06-12), use Receive Date: ensure the shipping date is not after the Receive Date. - All date output **must be after 2025-04-01**. If not, double-check format; if still not compliant, proceed as below. 5. **Ambiguity and Failsafes:** - If no explicit ship date is found, or ambiguity remains, output **only**: `Not found`. - Only output the date if unambiguous and correctly formatted. Otherwise, output `Not found`. 6. **Output:** - Output strictly one value: the converted shipping date in YYYY-MM-DD format, or exactly `Not found`. - **Do not include any spaces, notes, or explanations.** ## Output Format - Output must be **a single string value only**: either the shipping date in YYYY-MM-DD format or `Not found`. - No extra text, explanations, spaces, or reasoning. ## Reasoning and Conclusion Order **ORDER:** 1. Extract and analyze (reasoning) the required date from the document, validating constraints. 2. Output only the single shipping date value or "Not found" (conclusion). ## Example Inputs and Outputs **Example 1** PDF: Westburne packing slip, "Ship Date" field reads '25 JAN 12'. Receive Date: 2025-06-12 Output: 2025-01-12 **Example 2** PDF: WESCO packing slip, landscape, top-left reads '03-24-25'. Receive Date: 2025-06-12 Output: 2025-03-24 (If the result was before 2025-04-01, check again; if still earlier, and you can't find another result, output "Not found") **Example 3** PDF: Unable to find any ship date. Output: Not found **Example 4** PDF: Westburne, Ship Date is 'DEC 05' Receive Date: 2025-06-12 Output: 2025-12-05 (Note: Year assumed to be 2025 as per instruction.) **Example 5** PDF: '12/06' present, Receive Date: 2025-06-12 Output: 2025-06-12 (Disambiguate order so shipping date is not later than receive date.) ## Edge Cases and Considerations - Never guess or invent a date value. - Yearless dates default to 2025. - If a shipping date is not found or not after 2025-04-01, output "Not found". - Never output explanation, multiple values, or any other text. **REMINDER:** Output only one date value in YYYY-MM-DD format or "Not found". Never invent, guess, or explain. Follow all constraints, including after 2025-04-01 rule and single-value output.
        """
'''
