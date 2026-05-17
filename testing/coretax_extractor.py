import base64
import json
import io
import os
import requests

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

class CoreTaxExtractor:
    def __init__(self, token, cookie):
        self.base_url = "https://coretaxdjp.pajak.go.id"
        self.headers = {
            "authority": "coretaxdjp.pajak.go.id",
            "authorization": f"Bearer {token}",
            "content-type": "application/json",
            "languageid": "id-ID",
            "referer": f"{self.base_url}/e-invoice-portal/id-ID/input-tax",
            "cookie": cookie,
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
            "x-dgt-code": "7AcAAA==",
            "accept": "application/json, text/plain, */*",
            "accept-language": "id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7"
        }
        self.taxpayer_id = self._extract_taxpayer_id(token)

    def _extract_taxpayer_id(self, token):
        try:
            payload_b64 = token.split('.')[1]
            missing_padding = len(payload_b64) % 4
            if missing_padding: payload_b64 += '=' * (4 - missing_padding)
            payload = json.loads(base64.b64decode(payload_b64).decode('utf-8'))
            tid = payload.get("taxpayer_id")
            if tid:
                print(f"[+] AKTIF: Login sebagai {payload.get('full_name')} ({tid})")
                return tid
        except Exception: pass
        return None

    def get_invoice_details(self, invoice_number):
        if not self.taxpayer_id:
            print("[!] Token TIDAK VALID.")
            return

        print(f"\n[*] MENCARI FAKTUR: {invoice_number}")
        search_url = f"{self.base_url}/einvoiceportal/api/inputinvoice/list"
        
        payload = {
            "BuyerTaxpayerAggregateIdentifier": self.taxpayer_id,
            "TaxpayerAggregateIdentifier": self.taxpayer_id,
            "First": 0, "Rows": 10, "SortField": "TaxInvoiceDate", "SortOrder": -1,
            "LanguageId": "id-ID",
            "Filters": [
                {"PropertyName": "TaxInvoiceNumber", "Value": invoice_number, "MatchMode": "equals"}
            ]
        }

        try:
            response = requests.post(search_url, headers=self.headers, json=payload)
            if response.status_code != 200:
                print(f"[!] Server Error {response.status_code}: {response.text}")
                return

            res_json = response.json()
            data_list = res_json.get("Payload", {}).get("Data", [])

            if not data_list:
                print(f"[-] Faktur {invoice_number} tidak ditemukan di akun ini.")
                self._show_available_invoices(search_url, payload)
                return

            inv = data_list[0]
            print(f"[+] Ditemukan! Penjual: {inv['SellerTaxpayerName']}")
            
            # AMBIL PDF
            print("[*] Mengunduh detail barang...")
            dl_url = f"{self.base_url}/einvoiceportal/api/DownloadInvoice/download-invoice-document"
            dl_payload = {
                "EInvoiceRecordIdentifier": inv["RecordId"],
                "EInvoiceAggregateIdentifier": inv["AggregateIdentifier"],
                "DocumentAggregateIdentifier": inv["DocumentFormAggregateIdentifier"],
                "TaxpayerAggregateIdentifier": self.taxpayer_id,
                "LetterNumber": invoice_number,
                "EInvoiceMenuType": "Input",
                "TaxInvoiceStatus": "APPROVED"
            }

            dl_resp = requests.post(dl_url, headers=self.headers, json=dl_payload)
            dl_json = dl_resp.json()

            if not dl_json.get("IsSuccessful"):
                print("[!] Gagal ambil PDF:", dl_json.get("Message"))
                return

            pdf_bytes = base64.b64decode(dl_json["Content"])
            self._parse_pdf(pdf_bytes)

        except Exception as e:
            print(f"[!] Error: {e}")

    def _show_available_invoices(self, url, original_payload):
        print("\n[*] Menampilkan 5 faktur terakhir yang tersedia di akun Anda:")
        test_payload = original_payload.copy()
        test_payload["Filters"] = []
        try:
            resp = requests.post(url, headers=self.headers, json=test_payload)
            invoices = resp.json().get("Payload", {}).get("Data", [])
            for i, d in enumerate(invoices[:5]):
                print(f"    {i+1}. No: {d['TaxInvoiceNumber']} | Dari: {d['SellerTaxpayerName']}")
            print("\n[TIP] Gunakan salah satu nomor di atas untuk mencoba script ini.")
        except: pass

    def _parse_pdf(self, pdf_bytes):
        if not pdfplumber:
            print("[!] Library pdfplumber tidak ditemukan.")
            return
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = pdf.pages[0].extract_text()
            print("\n" + "="*60 + "\n" + text + "\n" + "="*60)

if __name__ == "__main__":
    # TOKEN BARU DARI HASIL DEBUG ANDA
    TOKEN_DEFAULT = "eyJhbGciOiJSUzI1NiIsImtpZCI6IjVENjlGREUxMTRBQzkzRTU1QzExQjMwQjIyMzFGQTRDMzlGNzNDMUJSUzI1NiIsInR5cCI6ImF0K2p3dCIsIng1dCI6IlhXbjk0UlNzay1WY0ViTUxJakg2VERuM1BCcyJ9.eyJuYmYiOjE3Nzg4MzM4NTEsImV4cCI6MTc3ODgzNTY1MSwiaXNzIjoiaHR0cHM6Ly9jb3JldGF4ZGpwLnBhamFrLmdvLmlkL2lkZW50aXR5cHJvdmlkZXJwb3J0YWwiLCJjbGllbnRfaWQiOiJjYXRzLXBvcnRhbC1hbmd1bGFyLWNsaWVudCIsInN1YiI6IjAwMzk5OTQ0MTMzMjIwMDAiLCJhdXRoX3RpbWUiOjE3Nzg4MzM4NDksImlkcCI6ImxvY2FsIiwiaWRlbnRpZmllciI6IjVlYjdkODM3LTcxMDYtZTU2Yi01NDc1LTVlNmU4NmU5MTBlNSIsInVzZXJfbmFtZSI6IjAwMzk5OTQ0MTMzMjIwMDAiLCJpbmFjdGl2aXR5VGltZW91dCI6IjkwMCIsImxhbmd1YWdlIjoiaWQtSUQiLCJsYXN0bG9naW50aW1lIjoiMjAyNi0wNS0xNVQxNTozMDo1MS4yNzIiLCJ0YXhvZmZpY2UiOiIzMjIiLCJmdWxsX25hbWUiOiJUSUdBU0FUVSBQRVJEQUdBTkdBTiBJTkRPTkVTSUEiLCJlbWFpbCI6IlBUVElHQVNBVFVQRVJEQUdBTkdBTklORE9ORVNJQUBHTUFJTC5DT00iLCJwZXJtaXNzaW9ucyI6Ils0LDE3LDE4LDIwLDIxLDI1LDI5LDM3LDQxLDQ1LDQ5LDUzLDU3LDYxLDYyLDYzLDY0LDY1LDkyLDkzLDk0LDk1LDk2LDk3LDk4LDk5LDkwMCw5ODAsOTg0LDQsMTcsMTgsMjAsMjEsMjUsMjksMzcsNDEsNDUsNjEsNjQsOTIsOTMsOTQsOTUsOTYsOTcsOTgsMTU4LDk4MCw5ODQsMTcsMTgsMjAsMjEsMjYsMzAsMzgsNDEsNDUsNTAsNTQsNTgsOTIsOTMsOTUsOTYsOTcsOTgsOTAwLDk4Miw5ODYsNCwxNywxOCwyMCwyMSwyNiwzMCwzNCwzOCw0MSw0NSw1MCw1NCw1OCw5Miw5Myw1LDk2LDk3LDk4LDk5LDE1OCw5MDAsOTgyLDk4Niw0LDE3LDIxLDI3LDMxLDM5LDQxLDQ1LDkyLDk1LDk2LDk3LDk4LDk4MSw0LDE3LDE4LDIwLDIxLDI3LDMxLDM1LDM5LDQxLDQ1LDUxLDU1LDU5LDkyLDkzLDk1LDk2LDk3LDk4LDk5LDkwMCw5ODEsOTg1LDQsMTcsMTgsMjAsOTIsOTYsMTI0LDE1OCwxOTksMjA0LDIwNSwyMDYsMjA3LDIwOCwyMDksMjEyLDIxMywyMjMsMjI0LDIyNiwyMjcsMjI4LDIyOSwyMzAsMjMxLDIzMiwyMzMsMjM0LDIzOCwyNDEsMjQ3LDI0OCwyNTMsMjU1LDI1NiwyNjAsMjYxLDI2MiwyNjMsMjY0LDI2OSwyNzAsMjcxLDI3MiwyNzQsMjc1LDI3NiwyNzcsMjc4LDI4MywyODQsMjg1LDc1MCw3NTIsNzUzLDc1NCw3NTUsNzU2LDc1Nyw3NjAsNzY2LDc2Nyw3NzIsNzc2LDc3OCw3NzksNzgxLDc4Miw3ODMsNzg0LDc4NSw3ODYsNzg5LDc5MCw3OTEsNzkzLDc5NCw3OTUsNzk2LDc5Nyw3OTgsODA0LDgwNSw4MDYsODA3LDgwOCw4MTIsODEzLDgxNCw4MTUsODE2LDgxNyw4MTgsODE5LDgyNyw4MzAsODMyLDgzMyw4MzQsODM1LDgzNiw4MzcsODQwLDg0Miw4NDMsODQ0LDg0NSw4NDYsODQ5LDEzMTAsNCwxNywxOCwyMCw5NiwxNTgsMjA0LDIwNSwyMDYsMjA3LDIwOCwyMDksMjEyLDIxMywyMTQsMjE1LDIyMywyMjQsMjI1LDIyNiwyMjcsMjI4LDIyOSwyMzAsMjMxLDIzMiwyMzMsMjM0LDIzNSwyMzYsMjM3LDIzOCwyMzksMjQwLDI0MSwyNDIsMjQzLDI0NCwyNDUsMjQ2LDI0NywyNDgsMjQ5LDI1MCwyNTEsMjUyLDI1MywyNTQsMjU1LDI2MCwyNjEsMjYyLDI2MywyNjQsMjY1LDI2NiwyNjcsMjY4LDI2OSwyNzAsMjcxLDI3MiwyNzMsMjc0LDI3NSwyNzYsMjc3LDI3OCwyNzksMjgwLDI4MSwyODIsMjgzLDI4NCwyODUsNzUwLDc1MSw3NTIsNzUzLDc1NCw3NTUsNzU2LDc1Nyw3NTgsNzU5LDc2MCw3NjEsNzYyLDc2Myw3NjQsNzY1LDc2Niw3NjcsNzY4LDc2OSw3NzAsNzcxLDc3Miw3NzMsNzc0LDc3NSw3NzYsNzc3LDc3OCw3NzksNzgwLDc4MSw3ODIsNzgzLDc4NCw3ODUsNzg2LDc4Nyw3ODgsNzg5LDc5MCw3OTEsNzkyLDc5Myw3OTQsNzk1LDc5Niw3OTcsNzk4LDc5OSw4MDAsODAxLDgwMiw4MDMsODA0LDgwNSw4MDYsODA3LDgwOCw4MDksODEwLDgxMSw4MTIsODEzLDgxNCw4MTUsODE2LDgxNyw4MTgsODE5LDgyMCw4MjEsODIyLDgyMyw4MjQsODI1LDgyNiw4MjcsODI4LDgyOSw4MzAsODMxLDgzMiw4MzMsODM0LDgzNSw4MzYsODM3LDgzOCw4MzksODQwLDg0MSw4NDIsODQzLDg0NCw4NDUsODQ2LDg0Nyw4NDgsODQ5LDEzMDAsMTMxMCw0LDE3LDIxLDI3LDMxLDM5LDQxLDkyLDkzLDk1LDk2LDk3LDk4LDk4MSw0LDE3LDE4LDIwLDIxLDI3LDMxLDM1LDM5LDQxLDQ1LDUxLDU1LDU5LDkyLDkzLDk1LDk2LDk3LDk4LDk5LDkwMCw5ODEsOTg1LDEzMzAsMTM0MCwxNywxOCwyMSwyNywzMSwzOSw0MSw0NSw5Miw5NCw5NSw5Niw5Nyw5OCw5ODEsNCwxNywxOCwyMCwyMSwyNywzMSwzOSw0MSw0NSw1MSw1NSw1OSw5Miw5Myw5NCw5NSw5Niw5Nyw5OCw5OSw5MDAsOTgxLDk4NSw0LDE3LDIxLDI4LDMyLDQwLDQxLDkyLDkzLDk1LDk2LDk3LDk4LDk4Myw0LDE3LDE4LDIwLDIxLDI4LDMyLDM2LDQwLDQxLDQ1LDUyLDU2LDYwLDkyLDkzLDk1LDk2LDk3LDk4LDk5LDkwMCw5ODMsOTg3LDEzMzUsMTM0NSw0LDEyNSwxMjYsMTI3LDEyOCwxMjksMTMwLDEzMSwxMzIsMTMzLDEzNCwxMzUsMTM2LDEzOCwxNDAsMTQyLDE2MywxNzAsMTcxLDE3MiwxNzMsMTc1LDE3OSwxODEsMTg3LDE4OCwxODksMTkwLDE5MSw2MDAsNjAxLDYwMyw2MDQsNjA1LDYxMSw2MTIsNjE0LDYxNSw2MTYsNjIyLDYyMyw2MjUsNjI2LDYyNyw2MzMsNjM0LDYzNiw2MzcsNjM4LDY0NCw2NDUsNjQ2LDk1MCw5NTMsOTU2LDk1OSw5NzcsNCwxOCwyMCwxMjUsMTI2LDEyNywxMjgsMTI5LDEzMCwxMzEsMTMyLDEzMywxMzQsMTM1LDEzNiwxMzcsMTM4LDEzOSwxNDAsMTQxLDE0MiwxNjMsMTY0LDE3MCwxNzEsMTcyLDE3MywxNzUsMTc5LDE4MSwxODcsMTg4LDE4OSwxOTAsMTkxLDYwMCw2MDEsNjAyLDYwMyw2MDQsNjA1LDYwNiw2MDcsNjA4LDYwOSw2MTAsNjExLDYxMiw2MTMsNjE0LDYxNSw2MTYsNjE3LDYxOCw2MTksNjIwLDYyMSw2MjIsNjIzLDYyNCw2MjUsNjI2LDYyNyw2MjgsNjI5LDYzMCw2MzEsNjMyLDYzMyw2MzQsNjM1LDYzNiw2MzcsNjM4LDYzOSw2NDAsNjQxLDY0Miw2NDMsNjQ0LDY0NSw2NDYsNjQ3LDY0OCw5NTAsOTUxLDk1Miw5NTMsOTU0LDk1NSw5NTYsOTU3LDk1OCw5NTksOTYwLDk2MSw5NzcsMTM0MCw0LDE0MywxNDQsMTQ1LDE0NiwxNDgsMTQ5LDE1MCwxNTEsMTUzLDE1NCwxNTUsMTU2LDE1OCwxNTksMTYwLDE2MSwxNjUsMTY2LDE2NywxNjgsMTc0LDE3NiwxNzcsMTc4LDE4MCwxODIsMTgzLDE4NCwxODUsMTg2LDY0OSw2NTAsNjUyLDY1Myw2NTQsNjU5LDY2MCw2NjIsNjYzLDY2NCw2NzAsNjcxLDY3Myw2NzQsNjc1LDY4MSw2ODIsNjg0LDY4NSw2ODYsNjkyLDY5Myw2OTUsNjk2LDY5Nyw5NjIsOTY1LDk2OCw5NzEsOTc0LDQsMTgsMjAsMTQzLDE0NCwxNDUsMTQ2LDE0NywxNDgsMTQ5LDE1MCwxNTEsMTUyLDE1MywxNTQsMTU1LDE1NiwxNTcsMTU4LDE1OSwxNjAsMTYxLDE2MiwxNjUsMTY2LDE2NywxNjgsMTY5LDE3NCwxNzYsMTc3LDE3OCwxODAsMTgyLDE4MywxODQsMTg1LDE4Niw2NDksNjUwLDY1MSw2NTIsNjUzLDY1NCw2NTUsNjU2LDY1Nyw2NTgsNjU5LDY2MCw2NjEsNjYyLDY2Myw2NjQsNjY1LDY2Niw2NjcsNjY4LDY2OSw2NzAsNjcxLDY3Miw2NzMsNjc0LDY3NSw2NzYsNjc3LDY3OCw2NzksNjgwLDY4MSw2ODIsNjgzLDY4NCw2ODUsNjg2LDY4Nyw2ODgsNjg5LDY5MCw2OTEsNjkyLDY5Myw2OTQsNjk1LDY5Niw2OTcsNjk4LDY5OSw3MDAsNzAxLDcwMiw5NjIsOTYzLDk2NCw5NjUsOTY2LDk2Nyw5NjgsOTY5LDk3MCw5NzEsOTcyLDk3Myw5NzQsOTc1LDk3Niw0LDE5OSwyMDAsMjAzLDIxMCwyMTEsMjE2LDIxNywyMTgsMjE5LDIyMCwyMjEsMjIyLDI4NiwyODcsMjg4LDI4OSwyOTAsMjkxLDI5MiwyOTMsMjk0LDI5NSwyOTYsMjk3LDI5OCwyOTksMzAwLDMwMSw0MTIsNDEzLDQxNCw0MTUsNDE2LDQxNyw0MTgsNDE5LDQyMCw0MjEsNDIyLDQyMyw0MjQsNDI1LDQyNiw0MjcsNDI4LDQyOSw0MzAsNDMxLDQzMiw0MzMsNDM0LDQzNSw0MzYsNDM3LDQzOCw0MzksNDQwLDQ0MSw0NDIsNDQzLDQ0NCw0NDUsNDQ2LDQ0Nyw0NDgsNDQ5LDQ1MCw0NTEsNDUyLDQ1Myw0NTQsNTAwLDUwMSw5MDgsOTIwLDksMTAsMTEsMTIsMTMsMTQsMTUsMTYsMTgsMjAsMjYsMzAsMzMsMzQsNTAsNTQsNTgsMTA3LDEwOCwxMDksMTE5LDEyMCwxMjEsMTIzLDEyNCwxOTIsMTkzLDE5NywxOTksMjAwLDIwNCwyMDUsMjA2LDIwNywyMDgsMjA5LDIxMCwyMTEsMjEyLDIxMywyMTYsMjE3LDIxOCwyMTksMjIwLDIyMSwyMjIsMjIzLDIyNCwyMjYsMjI3LDIyOCwyMjksMjMwLDIzMSwyMzIsMjMzLDIzNCwyMzgsMjQxLDI0NywyNDgsMjUzLDI1NSwyNTYsMjU3LDI1OCwyNjAsMjYxLDI2MiwyNjMsMjY0LDI2OSwyNzAsMjcxLDI3MiwyNzQsMjc1LDI3NiwyNzcsMjc4LDI4MywyODQsMjg1LDI4NiwyODcsMjg4LDI4OSwyOTAsMjkxLDI5MiwyOTMsMjk0LDI5NSwyOTYsMjk3LDI5OCwyOTksMzAwLDMwMSwzMTAsNDAwLDQwMSw0MTIsNDEzLDQxNCw0MTUsNDE2LDQxNyw0MTgsNDE5LDQyMCw0MjYsNDI3LDQyOCw0MjksNDMwLDQzMSw0MzIsNDMzLDQzNCw0MzUsNDM2LDQzNyw0MzgsNDM5LDQ0MCw0NDEsNDQyLDQ0Myw0NDQsNDQ1LDQ1Myw0NTQsNTAwLDUwMSw3NTAsNzUyLDc1Myw3NTQsNzU1LDc1Niw3NTcsNzYwLDc2Niw3NjcsNzcyLDc3Niw3NzgsNzc5LDc4MSw3ODIsNzgzLDc4NCw3ODUsNzg2LDc4OSw3OTAsNzkxLDc5Myw3OTQsNzk1LDc5Niw3OTcsNzk4LDgwNCw4MDUsODA2LDgwNyw4MDgsODE2LDgxNyw4MTgsODE5LDgyNyw4MzAsODMyLDgzMyw4MzQsODM1LDgzNiw4MzcsODQwLDg0Miw4NDMsODQ0LDg0NSw4NDYsODQ5LDkwNSw5MDYsOTA3LDkwOCw5MjAsOTIxLDkzNyw5MzgsOTgyLDk4NiwxMzEwLDQsMTk4LDE5OSwyMDAsMjAzLDIwNiwyMTAsMjExLDIxNiwyMTcsMjE4LDIxOSwyMjAsMjE0LDgxNSwyMjEsMjIyLDI1NywyNTgsMjg2LDI4NywyODgsMjg5LDI5MCwyOTEsMjkyLDI5MywyOTQsMjk1LDI5NiwyOTcsMjk4LDI5OSwzMDAsMzAxLDMxMCw0MDEsNDEyLDQxMyw0MTQsNDE1LDQxNiw0MTcsNDE4LDQxOSw0MjYsNDI3LDQyOCw0MjksNDMwLDQzMSw0MzIsNDMzLDQzNCw0MzUsNDM2LDQzNyw0MzgsNDM5LDQ0MCw0NDEsNDQyLDQ0Myw0NDQsNDQ1LDQ1Myw0NTQsNTAwLDUwMSw5MDgsOTIwLDkyMV0iLCJyb2xlcyI6IlszMywzNCw0MCw0MSwzMCwzMiwzOCwzOSw0NiwzMSwzNyw0NCw0NSwxMDAyLDM1LDM2LDQyLDQzLDQ3LDExMDQsMTIwOV0iLCJ1c2VyX3JvbGVzIjoiWzMxLDExMDQsMTIwOSw0NCwxMSwzNCw0MCwzNiwzOCwzMl0iLCJ0YXhwYXllcl9pZCI6IyIyNzViYjA3YS1iY2EzLTQwNzItYWYxNS02ZGFjMjdhNjcwOGUiLCJkYXRhYWNjZXNzcG9saWNpZXMiOiJbXCJWQVRfVkFUVE5URVZBVENcIixcIkFsd2F5c0FsbG93VGF4cGF5ZXJFZGl0XCIsXCJWQVRfRFZBVFwiLFwiQWxsb3dFZGl0SWZUYXhwYXllclRheFJlZ2lvbk1hdGNoZXNVc2VyVGF4UmVnaW9uXCIsXCJWQVRfVkFUXCIsXCJBbGxvd0VkaXRJZlRheHBheWVyVGF4T2ZmaWNlTWF0Y2hlc1VzZXJUYXhPZmZpY2VcIl0iLCJyZXByZXNlbnRhdGl2ZVRheFR5cGVzIjoiW10iLCJyZXByZXNlbnRhdGl2ZVN1YnNlcnZpY2VzIjoiW10iLCJyZXByZXNlbnRhdGl2ZUZ1bGxBZGRyZXNzIjoiUEVSVU0gQUxJVklBIFJFU0lERU5DRSBMSyBJIEQgTk8uNSwgIFJUIDAwMSwgIFJXIDAwMCwgU1VNQkVSIEFHVU5HLCBLRU1JTElORywgS09UQSBCQU5EQVIgTEFNUFVORywgTEFNUFVORywgSW5kb25lc2lhIDM1MTU5IiwicmVwcmVzZW50YXRpdmVQaG9uZSI6IjA4MjIwMjE0NjU3IiwicmVwcmVzZW50YXRpdmVFbWFpbEFkZHJlc3MiOiJhaWt1cm5pYXdhbjc4OUBnbWFpbC5jb20iLCJyZXByZXNlbnRhdGl2ZVR5cGUiOiJpbnRlcm5hbCIsIkltcGVyc29uYXRpbmciOiJ0cnVlIiwiUmVwcmVzZW50YXRpdmVJZCI6ImZkYWZkMzlmLTJmY2MtNGMxZi1mNTE3LWU0NTVmMzQzNjBkMiIsIlJlcHJlc2VudGF0aXZlVXNlcklkIjoiNzYwNTdmOTEtMzRjZi05YmZiLTQzM2UtMDIzOTg4MTMzYTcwIiwiUmVwcmVzZW50YXRpdmVVc2VybmFtZSI6IkFDSE1BRCBJTkRSQSBLVVJOSUFXQU4iLCJSZXByZXNlbnRhdGl2ZVRpbiI6IjE4NzEwOTAxMDc4OTAwNzMiLCJqdGkiOiI5RDU1MUNENUExQzNCQkZENkNFNUZCRkVBODAyQzA3OCIsInNpZCI6IjM2RTVDMDE4NEM0QjFEOTEyNzlEOUE0RTI2REFCQUNFIiwiaWF0IjoxNzc4ODMzODUxLCJzY29wZSI6WyJvcGVuaWQiLCJwcm9maWxlIiwiY2F0c2hvbWVwb3J0YWwtYXBpIl0sImFtciI6WyJwd2QiXX0.o5s7g1RvY183zeWhdVoJnPakcW7ZqVSqe2lxI1FANEVA_PAwIMkP28YbVq5jwM-tbtevqnKs11hbV-U4b0XMuaHNLIob80kkvFQUOhqHWEorV1kvl8FfjszdIQE5ljfSiMr0OqlcehKW2sKb37Nav0BRPnCXGJeJBXvqmCtM8wbfqrHTAXwHd7hi2tkZ9uZhfeOQ2wFYnJbWdgr1CPnE_sSEeqZ64gBeQwmLioi9aOPnLoUA3S1pUc0JYdVrrTWEJrRjQ5Nx1ZtKVf7TymL-d_OagncvTdJpY10FIum5wWNROVuKCU3kFe6QuPUdkfO6ATbGNwmTEPoVVUMyFdrWmqxojJtH-0ppmzMKzaAiUvRbTXvRdeuXokXyQLtnPdGZhaOUxeQ6MUY1VXMDlIaiy9C_J7_K0mIINByodOzQyzgkKgDsPIP8ZUkfx-m2YkqkeLV0u1LWk5NvzhzNV_Txk9n68o7VTfvL2E3YPFQfTHAcZ5ms65oPYa8ZnXPQYrf7q9xaKoRbIjyLwPXBGj8rj01HOLEiakzxqCbFurnYPKld9dnRSYDjkoQAR2mAw59WDWU9vpFXsXgsCFHyd2PpxJggUWX88ZGU3QnVzE2KE5TXMNTXyqWjfCaWjZNIEvgb3SOWehRlbV60b6NKSrkiFNg87-5OgBHm1L8thMoQQNE"
    
    # COOKIE BARU DARI HASIL DEBUG ANDA
    COOKIE_DEFAULT = "Portal-xsrf-cookie=CfDJ8LBptqruC79FgZVBt06-M5w0aiF2W9zWeTriwQ-cEH8oqEO5_CD7SkKsES3XJ5MbxiHKOxMzFb25kFl6pqnL5z6lrowTiO21fkPTnSMVhLLHmP_aF2ewPltAX3U67n4yRJNoRFK6fkbklFPMCh_QWEs; idsrv.session=36E5C0184C4B1D91279D9A4E26DABACE; Portal=CfDJ8LBptqruC79FgZVBt06-M5zIokHv-K-Lou2sNX-RirCTfJIqm1ITepnPQHMt_N16gFhSHo6i4qgDaI5J_xP7cC94S4Yv1U_Qnr7X56zA7aSuy2w9emqOjXqca-grQC0kP1DKpSSWmV4XZXozijroaDsplN2dWpj1cmuR5YaD_XmEoO4wVEQIUkOF5f1qFtT1T6bbYEBFLSKmLABkNNG4ijRXh11b-K5yJgZTxYWzjRyDORGhCwyeHxZhqHNdJS0Ba0xEdezJ4F2U3gVlbHoo4PxCAaoGZzylLwFBDgnjGqatjryUP9Ph_FArxg22FV514dp7gdB3AvktVlu-kjiVYVntJ-3Bm6cFxHcpPai9aUWiEotHJlyjAWitw_c4XA9Ozt7nsZ9vsFVMTG-GqldxmzQ4ehYjXsW03mJScg8pG0wQgsVF3WZ7wvFRHyRJqCGBc-xv7yrDyRqK4fWLYhmyPKrpFDtYwT_SurQCtE3c1_uIxohu_nC31F5EjY7AkQvHgSMc1oMTE3Gh-9Q7sFu4RlIK4VX0ZPw-LePHnHq3TvrjVDIO9BrGRd0UHbmChSu4VffIHfrsne9-xhlljuj3PoJhuglFDbSyalX1-KKgnkiO8mgrn1BWOAibeV8MspQ_33n9ZVzdp0vtVJmtve8EplZ7y1TwEhEUQgb8KMfjKarGn49fcclsi1TKBnRFyKbeYczJ3jL_fKRYsYO_GaRX_56LAmW4cOjDngcyL6PRth7PY6kf0R7XKIaRomd8WXZ268WLPrEIRPvbbevbb90IOoFJKA-oz_1NGl15cGcFdaDuvmRS7javiHPIZrksg92DzH3ZbIwBY77kOl1GQRieC4jtJuYuOjCthnwIhQexVBpMdmZ-DL1GHfJA6pfmnciHzBhAS1JISdtQgjPFV0CYagzSh-PguZahY8HvShGZWbGU5qeCHSkaTv2VYiQAnP2cCZr7HIEN32KJa_Pmn37-C0FkK0QrSNCzFbRv3fLznwxI; selectedLanguage=id-ID; sl-session=rzdYTzYfCGpRr5Up+2EcQg=="

    print("=== CORETAX EXTRACTOR V4 (LIVE DEBUG) ===")
    token_input = input(f"Masukkan Bearer Token [Kosongkan untuk pakai token terbaru]: ").strip()
    token = token_input if token_input else TOKEN_DEFAULT
    
    # Faktur yang ingin dicari (sesuaikan jika ingin cari yang lain)
    no_faktur = "04002600169909529"
    
    extractor = CoreTaxExtractor(token, COOKIE_DEFAULT)
    extractor.get_invoice_details(no_faktur)
