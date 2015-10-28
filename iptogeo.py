import urllib2
import json
import xlrd
import xlwt


class IpToGeo(object):

    """read excel file for ip and get it geo data from https://www.iplocation.net/"""

    def __init__(self):
        super(IpToGeo, self).__init__()
        self.out = []

        self.ips = self.read_data("ips.xls")

        self.url = "https://ipinfo.io/"

    def read_data(self, spreadsheet):
        workbook = xlrd.open_workbook(spreadsheet)
        worksheet = workbook.sheet_by_name("Sheet1")
        num_rows = worksheet.nrows

        ips = []
        for row in range(0, num_rows):
            value = worksheet.cell_value(row, 0)
            ips.append(value)
        return ips

    def loop_ips(self):
        for ip in self.ips:
            self.out.append(self.get_data(ip))
        self.write_data(self.out)

    def get_data(self, ip):
        response = urllib2.urlopen(
            self.url + ip + "/json?token=iplocation.net").read()
        data = json.loads(response)
        city = data[u"city"]
        region = data[u'region']
        print(u"""
city:\t{}
state:\t{}
            """.format(city, region))
        return ip, city, region

    def write_data(self, data):
        xls_file = xlwt.Workbook()
        xls_file_name = "out_ips.xls"
        ips_sheet = xls_file.add_sheet("out")
        for index_row, row in enumerate(data):
            for index_item, item in enumerate(row):
                ips_sheet.write(index_row, index_item, item)
        xls_file.save(xls_file_name)


if __name__ == '__main__':
    togeo = IpToGeo()
    togeo.loop_ips()
