import pandas as pd
from playwright.sync_api import sync_playwright

def run(playwright):
    try:
        data = pd.read_excel(r"C:\Users\rajac\OneDrive\Desktop\RegressionTesting.xlsx")
        print('Excel Opened')
    except Exception as e:
        print(f"Failed to read Excel file: {e}")
        return
    
    gender_map = {'m': 'Male', 'f': 'Female', 'p': 'Prefer not to say'}
    try:
        with playwright.chromium.launch(headless=False) as browser:
            page = browser.new_page()

            for index, row in data.iterrows():
                page.goto("https://docs.google.com/forms/d/e/1FAIpQLSf-FvGRLUQ4XkHZyp2aw1qz2GpAcA_dDz8OqgcP780mjJ6Esg/viewform?vc=0&c=0&w=1&flr=0")
                print(f'Step1 done for row {index}')
                
                try:
                    name_input = page.locator('input[aria-labelledby="i1"][aria-describedby="i2 i3"]')
                    family_name_input = page.locator('input[aria-labelledby="i5"][aria-describedby="i6 i7"]')
                    address_input = page.locator('input[aria-labelledby="i9"][aria-describedby="i10 i11"]')
                    operations = page.locator('#i21')
                    finance = page.locator('#i24')
                    age_input = page.locator('input[aria-labelledby="i27"][aria-describedby="i28 i29"]')


                    name_input.fill(row['Name'])
                    page.wait_for_timeout(2000)
                    family_name_input.fill(row['Family Name'])
                    page.wait_for_timeout(2000)
                    address_input.fill(row['Address'])
                    page.wait_for_timeout(2000)

                    

                    page.click('span.vRMGwf.oJeWuf')
                    page.wait_for_timeout(2000)

                    # Smooth scroll to a specific position
                    page.evaluate("""
                            window.scrollTo({
                            top: 1000,
                            left: 0,
                            behavior: 'smooth'
                            });
                            """)

                    gender_option_text = gender_map.get(row['Gender'].lower(), 'Prefer not to say')
                    print(f"Trying to click on: {gender_option_text}")
                    gender_option_selector = f"div[role='option'][data-value='{gender_option_text}']"
                    page.wait_for_selector(gender_option_selector, state="visible", timeout=2000).click()

                    page.wait_for_timeout(2000)
                    
                    if row['Department']=="Operations":
                        operations.click()
                    elif row['Department']=="Finance":
                        finance.click()
                    else:
                        "Error in department"

                    page.wait_for_timeout(2000)
                    age_input.click()
                    age_input.fill(str(row['Age']))

                    submit_button = page.locator('span.NPEfkd.RveJvd.snByac:text("Submit")')
                    submit_button.click()
                    page.wait_for_timeout(2000)
                    page.wait_for_load_state('networkidle')
                except Exception as e:
                    print(f"Failed to process form for row {index}: {e}")

    except Exception as e:
        print(f"Failed to launch browser or load page: {e}")

with sync_playwright() as playwright:
    run(playwright)
