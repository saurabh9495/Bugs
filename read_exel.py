import xlrd
import requests

# Give the location of the file
loc = ("intel.com-202006192010.xlsx")

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# For row 0 and column 0
print(sheet.cell_value(0, 0))

# Extracting number of rows
print(sheet.nrows)
nor = sheet.nrows

for i in range(0, nor):
    host = sheet.cell_value(i, 0)
    print(i, "--", host)
    http_host = "http://"+host
    https_host = "https://"+host
    try:
        r = requests.get(http_host)
        # print(r.headers)
        # print(r.text)
        if r.status_code < 200:
            i = open("100_informational_response.txt", "a+")
            i.write("\n{}".format(http_host))
            i.close()
        elif r.status_code >= 200 and r.status_code < 300:
            j = open("200_successful.txt", "a+")
            j.write("\n{}".format(http_host))
            j.close()
        elif r.status_code >= 300 and r.status_code < 400:
            k = open("300_redirection.txt", "a+")
            k.write("\n{}".format(http_host))
            k.close()
        elif r.status_code >= 400 and r.status_code < 500:
            l = open("400_client_error.txt", "a+")
            l.write("\n{}".format(http_host))
            l.close()
        else:
            m = open("500_server_error.txt", "a+")
            m.write("\n{}".format(http_host))
            m.close()
    except Exception as e:
        print(e)

    try:
        s = requests.get(https_host, verify=False)
        # print(s.headers)
        # print(s.text)
        if s.status_code < 200:
            n = open("100_informational_response.txt", "a+")
            n.write("\n{}".format(https_host))
            n.write("\n")
            n.close()
        elif s.status_code >= 200 and s.status_code < 300:
            o = open("200_successful.txt", "a+")
            o.write("\n{}".format(https_host))
            o.write("\n")
            o.close()
        elif s.status_code >= 300 and s.status_code < 400:
            p = open("300_redirection.txt", "a+")
            p.write("\n{}".format(https_host))
            p.write("\n")
            p.close()
        elif s.status_code >= 400 and s.status_code < 500:
            q = open("400_client_error.txt", "a+")
            q.write("\n{}".format(https_host))
            q.write("\n")
            q.close()
        else:
            t = open("500_server_error.txt", "a+")
            t.write("\n{}".format(https_host))
            t.write("\n")
            t.close()
    except Exception as e:
        print(e)
