import xlsxwriter


class Person():
    def __init__(self, name, unit, date, place, objename):
        self.name = name
        self.unit = unit
        self.date = date
        self.place = place
        self.objename = objename


workbook = xlsxwriter.Workbook('Expenses01.xlsx')
worksheet = workbook.add_worksheet()

y = int(input("Lütfen insan sayısı giriniz: "))

obje = []
human = Person(name="aa", unit="bb", date="cc", place="kk", objename="nn")
for m in range(y):
    name = input("İsim: ")
    unit = input("Ünite: ")
    date = input("Tarih: ")
    place = input("Yer: ")
    objename = input("Obje İsmi: ")
    human = Person(name, unit, date, place, objename)
    
    obje.append(human)


cell_format = workbook.add_format({'bg_color': '#98FB98'})
header = ['Talebi oluşturan kişinin adı ve soyadını giriniz.', 'Talebi Oluşturan Kişinin Bulunduğu Birim:', 'Talebin Oluşturulma Tarihi:', 'Talep edilen ürünün kullanım yeri:', "Talep Edilen Ürünün Adı:"]
worksheet.write_row(0, 0, header, cell_format)

for m, obje in enumerate(obje):
    expenses = [
        [obje.name, obje.unit, obje.date, obje.place, obje.objename],
    ]

    col = 0

    for i, item in enumerate(expenses):
        
        # Farklı renkler için renk listesi
        colors = ['#FFC0CB', '#FFE4E1', '#F0E68C', '#98FB98']

        # Satırın üzerine yazmak için stil oluşturma
        cell_format = workbook.add_format({'bg_color': colors[i]})

        # Satırı yazdırma
        worksheet.write_row(m + 1, col, item, cell_format)

workbook.close()