import gspread
from gspread_formatting import *
from oauth2client.service_account import ServiceAccountCredentials
import json
import requests
import time
import os

def check_package_status():

    scope = [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/drive.file'
    ]
    file_name = 'client_key.json'
    creds = ServiceAccountCredentials.from_json_keyfile_name(file_name,scope)
    client = gspread.authorize(creds)

    sheet = client.open('Shipments').sheet1 #Takes up to 5 seconds to open ( maybe leave it open )
    expected_headers = ['NAME', 'Shipping country', 'Destination COUNTRY (CITY)', 'LABEL CREATED DATE', 'Pickup / Drop off date', 'LABEL NUMBER', 'DESCRIPTION', 'STATUS', 'People Working', 'Instructions( for bot )']

    python_sheet = sheet.get_all_records(expected_headers=expected_headers)

    col = sheet.col_values(10)
    to_check = [i+1 for i, x in enumerate(col) if x == 'CHECK']
    result = []

    for i in range(len(to_check)):
        cell = sheet.cell(to_check[i],6).value
        result.append(cell)


    directory = os.path.dirname(__file__)
    file = os.path.join(directory, "getstatus.json")

    cookies = {
        'CONSENTMGR': 'c1:0%7Cc3:0%7Cc6:0%7Cc7:0%7Cc8:0%7Cc9:1%7Cts:1715353141812%7Cconsent:true%7Cc14:1',
        'ups_remember_me': '|mcb|',
        'ups_language_preference': 'en_FR',
        '_ga': 'GA1.1.1021663641.1715681286',
        '_ga_RHB0YE98J8': 'GS1.1.1715681285.1.1.1715681359.0.0.0',
        'upsShipPref': 'ews',
        'jedi_loc': 'eyJibG9ja2VkIjp0cnVlfQ%3D%3D',
        'sharedsession': '0a96886b-a8b1-428b-907c-58680b617c80:w',
        'X-CSRF-TOKEN': 'CfDJ8Jcj9GhlwkdBikuRYzfhrpL13raw5jyfjdM2B15XEA9LFNrjoOxDuLWfBBICJ_SB_yt-vNI3bUqBinw3s5-74TOFjfTUoETQqeU3HGf2po-aY3jOfFP6rGGtmxMXhQabF7I8ihbhsNm5EkDQD-jXW0A',
        'PIM-SESSION-ID': 'Rs3vbHtHPL7ffOaX',
        'st_cur_page': 'st_track',
        'dtCookie': 'v_4_srv_1_sn_D12A514BD509F233725EC5933A0947F3_perc_32485_ol_0_mul_3_app-3A83374a5b8f461e02_0',
        'AKA_A2': 'A',
        '_abck': '89682D3396F0F7F95C73C6DC3978C121~0~YAAQhOJIFwfxBLuPAQAAyKmX8Qzfi+emwz9jO04dfdlr/TItdCtq9Kav6X7lMB60DnE1YMcg7lEKX1G1WwCMpFh3fahywBYj3eHtjL1b1gfLmtJUyzp29/W3HGkOfejV74MsCN9Kzy2C4PkzWoCrhnS/NBZvE/5/0AwFCS9uICxTLa0gMREQKibYBnDiQZBKsX3F25gOu+p7l+qzTZmrNfaBe/4ratjy6zCe4a50rruRNRRqZKZDt5my8DtZgNUohZan1SKYI5DCrsdh+BtDVJYw0D+2BVUTimSzPT3ad1EoLEFOlWQDnR4UYR7NL+C3gsxE6O+SrlR/1yRZM2XuoAwGjyQVgBa9SA1egvu7nx7DM101KtZAW2S4JBqIZPlbhSf2tassFt7CRCjY8tDAfzvStSDs~-1~-1~1717748799',
        'bm_ss': 'ab8e18ef4e',
        'ak_bmsc': 'C283D1B6730F5DEFC9D5C950D00F5411~000000000000000000000000000000~YAAQydhLFzb+MMOPAQAARD6Y8RgFqMjZLDq3wEZyeZKWvIRniSSsbH+gvxyR5htv3zHnGcTohyo6t1uwU+kRA9EWznliz9FPTqG2D/38UuryprgVdGp4MEBxzjVrBUW7v3Ibq3/8VmD1Ey5VQ5TPkvIGCXXE3E5RKRb8jANNucw4AI9tpfvUW+55w5Kg3iDZm8Zrw6gDhe9y7K69wFw2xEVCaejFpeSX+jNE5IQg6U6yvxK7CMWpGKCa18w8TC/XBjXVPPa716I6I75HZKPkfw5bMzILkWUU1sPwmHNC6QBIboYSM2jpiXjoqr0ycpJyreUzTa2NYhJ5sNCLAB+T8+U+knHPr/m+BuKJZapxqMJ7ODbfPwXK8mNlesZHAAEm58mlBjOOhw==',
        'X-XSRF-TOKEN-ST': 'CfDJ8Jcj9GhlwkdBikuRYzfhrpLVb9cBSE87bW4N4mmypC7aGBcZu9wqiW4HWUzZwmkwrCnX76jr15nOteQuw9PVMJang1S7N2dtG6Bo9dQ26Fv4KazRdc4TxQBtWP3exXkLvR_SbJ9rZOqxNeG1ktMu6_A',
        'bm_s': 'YAAQhOJIF+z3BLuPAQAAUO+Y8QENurkW5e3Yd2ICxfvzDrEKeQexEsQ2nzmjCxJ9ey50CILJouaxldf1ITj+TTf03JNG2mFRI81mw56L/HdkP87krwfss7cQ2x8GekHvoF0MGY9QG97uUdJzkePwadTQEzw9ujHqCug/i9BQNo4kyF8IuW2zRxro1QSvqyiw1KWLOlEV1BpvxJxyESPBYn98RswJyPwzU1jk5xYt2Hpyy2fPRxyvp9Yvc14GqjLBVybyTGPGb6+xQ1yietsiM5GsqOq9EEVTVPfeb7NJjHD2UKTgewLrIY+rE6VvVKF0tjPvJFtGLRrZ1JCHzhVszc8=',
        'bm_sz': 'A89A7BDB68A3C069060D700C70F2E311~YAAQhOJIF+33BLuPAQAAUO+Y8RitXaaIbhM/2P/OmIRGlJeJbGGCGbwfKSyLOYsCVqDurJ17usx3aN2xSufdjNhMTElfefhZlasFLzdLvCv8jvgdav5ZeaiDx7EeT76EFjBhKSEczKA45XoDdsWdaxvV87gHuzBBvypcf7fI+ZzZo4v0Gu5fho/MjWhslR/jWXUUQezcH9lmTjs9ZaFgTtHOdwtJRF1z0TWF6Ls42G3nk5xo7TuAZ0Mo5jheP6UmfXK4HVo3wx2kgv5mUjV3kVqMz7/RMI3l1gQZq5YqfSKaLl9G00VaixE2CJ/wNPxJs0hYpkqF7GD2XMnnxDO4iFli7NuZOhquFBCdbiuLnDFCwh7W45kjBK9jlcslkwT5iP/vveQBJYwFY5m8ijA9gyShoSHpCxGoeHmsyDg=~3224645~3290677',
        'utag_main': 'v_id:018f6303ad6e001689773906344d05075003b06d00b3b$_sn:28$_se:5$_ss:0$_st:1717747083140$ses_id:1717745199302%3Bexp-session$_pn:3%3Bexp-session$fs_sample_user:false%3Bexp-session',
        'RT': '"z=1&dm=ups.com&si=78304914-e658-4830-8fee-294e42cfc7b2&ss=lx4d51s7&sl=2&tt=n9&bcn=%2F%2F02179915.akstat.io%2F"',
        'bm_sv': '3C3E0C81336C13D70917835BE9CF65F2~YAAQydhLF7knMcOPAQAANfSY8RjrFJTujJ9sVFtA8gPmj7orA7so+Usp/1k5Gfx4JxgSlWbD/wNSvn5lr+YzJ1w2cdiZZnzJIvoo5gW9swZHHMG8c0lbG3hAITX/MDUJKt9AXJOZm7Vbe6/W9NVTJyJMDv5uUDiMj4v9fn/Dy6jAu3fD5WBrXWc1HATV5mJAfZQPwseJEAcz0yuLDtUlgJ2tSJhQUKyzF9C8vf4Q0HmjfYQL7f57BCux346z~1',
    }

    headers = {
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'en-GB,en-US;q=0.9,en;q=0.8',
        'content-type': 'application/json',
        'origin': 'https://www.ups.com',
        'priority': 'u=1, i',
        'sec-ch-ua': '"Google Chrome";v="125", "Chromium";v="125", "Not.A/Brand";v="24"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"macOS"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-site',
        'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
        'x-xsrf-token': 'CfDJ8Jcj9GhlwkdBikuRYzfhrpLVb9cBSE87bW4N4mmypC7aGBcZu9wqiW4HWUzZwmkwrCnX76jr15nOteQuw9PVMJang1S7N2dtG6Bo9dQ26Fv4KazRdc4TxQBtWP3exXkLvR_SbJ9rZOqxNeG1ktMu6_A',
    }

    params = {
        'loc': 'en_FR',
    }

    for a in range(len(to_check)):

        json_data = {
            'Locale': 'en_FR',
            'TrackingNumber': [
                f"{result[a]}",
            ],
            'Requester': 'quic',
            'returnToValue': '',
        }
        response = requests.post(
            'https://webapis.ups.com/track/api/Track/GetStatus',
            params=params,
            cookies=cookies,
            headers=headers,
            json=json_data,
        )
        if response.status_code == 200:
            try:
                with open(file, 'w') as f:
                    reponse_json = response.json()
                    json.dump(reponse_json, f, indent=4)
            except:
                pass
        else:
            print("Failed with code :", response.status_code)

        with open(file, 'r') as infile:
            data = json.load(infile)

        package_status = data['trackDetails'][0]['packageStatus']
        progress = data['trackDetails'][0]['progressBarType']

        if "Exception" in progress or "Returned" in progress:
            sheet.format(f"A{to_check[a]}:I{to_check[a]}", {"backgroundColor": { "red": 1, "green": 0.5, "blue": 0}})
            sheet.update_cell(to_check[a],10,'OK')
        elif "Delivered" in progress:
            sheet.format(f"A{to_check[a]}:I{to_check[a]}", {"backgroundColor": { "red": 0.41, "green": 0.65, "blue": 0.3}})
            sheet.update_cell(to_check[a],10,'OK')
            sheet.update_cell(to_check[a],8,f"Delivered")
        else:
            sheet.format(f"A{to_check[a]}:I{to_check[a]}", {"backgroundColor": { "red": 1, "green": 1, "blue": 0}})
            sheet.update_cell(to_check[a],8,f"{package_status}")
        time.sleep(1)
    print("END")

if __name__ == "__main__":
    while True:
        check_package_status()
        time.sleep(900)  # Sleep for 15 minutes
