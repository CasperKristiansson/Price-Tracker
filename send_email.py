import smtplib
import pandas
import login

def send_mail(product_url, companies, recent_price, procent_change):
    gmail_user = login.EMAIL
    gmail_password = login.PASSWORD
    to = [login.TARGET_EMAIL]

    subject = 'Product Price Update Logitech z920s'
    body = f'''
{companies[0]}: {recent_price[0]}kr - {procent_change[0]}%
{product_url[0]}

{companies[1]}: {recent_price[1]}kr - {procent_change[1]}%
{product_url[1]}

{companies[2]}: {recent_price[2]}kr - {procent_change[2]}%
{product_url[2]}

{companies[3]}: {recent_price[3]}kr - {procent_change[3]}%
{product_url[3]}

{companies[4]}: {recent_price[4]}kr - {procent_change[4]}%
{product_url[4]}
    '''

    email_text = """\
From: %s
To: %s
Subject: %s
%s
""" % (gmail_user, ", ".join(to), subject, body)

    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(gmail_user, gmail_password)
    server.sendmail(gmail_user, to, email_text)
    server.close()

    print("Mail is successfully sent")

def data():
    product_url = []
    lowest_price = []
    recent_price = []
    i = 0

    excel_data_df = pandas.read_excel('Data.xlsx', sheet_name='Sheet1')
    while i < len(excel_data_df.columns):
        column = excel_data_df.columns[i]
        selected_data = excel_data_df[column].tolist()

        product_url.append(selected_data[0])
        min_price = selected_data[-1]
        recent_price.append(selected_data[-1])

        for price in selected_data:
            if str(price).isdigit():
                try:
                    if int(price) < int(min_price):
                        min_price = price
                except Exception as e:
                    print(e)

        lowest_price.append(min_price)
        i += 1

    return lowest_price, product_url, excel_data_df.columns.ravel(), recent_price

def compare_price(product_price, lowest_price):
    procent_change = []
    for current_price, old_price in zip(product_price, lowest_price):
        try:
            procent_change.append(round((int(current_price) / int(old_price)) * 100))
        except:
            procent_change.append("null")

    return procent_change

def main():
    LOWEST_PRICE, PRODUCT_URL, COMPANIES, RECENT_PRICE = data()
    PROCENT_CHANGE = compare_price(RECENT_PRICE, LOWEST_PRICE)
    send_mail(PRODUCT_URL, COMPANIES, RECENT_PRICE, PROCENT_CHANGE)


if __name__ == '__main__':
    main()
