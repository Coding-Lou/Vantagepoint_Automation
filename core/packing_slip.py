from openai import OpenAI
import pandas as pd
import os

def main(fileName, receiveDate):
    try:
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise ValueError("OPENAI_API_KEY environment variable is not set")
        client = OpenAI(api_key=api_key)

        file = client.files.create(
            file=open(f"packing_slip/{fileName}", "rb"),
            purpose="user_data"
        )

        prompt_text = """
You are an expert OCR and document-understanding system. You will be given a PDF containing shipment information. Find the date directly below to the label "Ship Date" (case-insensitive). The date follows the format YY MMM DD (e.g., 25 JAN 12). Convert it to YYYY-MM-DD: interpret YY as 20YY, map month abbreviation to two digits, keep the day. Output only the converted date in YYYY-MM-DD. If missing or unreadable, output null. Examples: 25 JAN 12 to 2025-01-12, 24 DEC 03 to 2024-12-03. Final output must be only one date value in YYYY-MM-DD format.
    """
        response = client.responses.create(
            model="gpt-4.1",
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
        You are a precise OCR and date extraction system. Auto-rotate the PDF to upright orientation; extract all visible text exactly as it appears.
Identify exactly three dates: Invoice Date or Order Date, Due Date, and Shipping Date or Delivery Date. Dates may appear in any common format (DD/MM, MM/DD, DD/MM/YYYY, MM/DD/YYYY, YY/MM/DD, DD.MM.YYYY, MM-DD-YY, DD-MM-YYYY). 
The receive date is {receiveDate} in YYYY-MM-DD format.
Rules:
1. Only use dates explicitly present in the document; do NOT invent, estimate, or calculate any dates.
2. If a date does not include a year, assume the year is 2025.
3. Treat all dates on the same page as using the same format.
4. Shipping Date must be no later than the Receive Date. Use this constraint to disambiguate dates with ambiguous day/month order. For example, if Receive Date is 2025-06-12 and the PDF shows 06/12, interpret it as 2025-06-12 (not 2025-12-06).
5. Convert all dates strictly to YYYY-MM-DD.
6. Output exactly the ShippingDate, with no spaces, explanations, or reasoning.
7. If any date cannot be determined unambiguously, output exactly "Not found".
The output MUST be STRICTLY to ONE DATE VALUE in YYYY-MM-DD or "Not found"
        """
'''
