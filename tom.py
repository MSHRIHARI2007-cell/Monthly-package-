from openpyxl import Workbook 

def to_ask_items():
    History = []
    while True:
        print("///******************HISTORY*********************///")
        name = input("Item name (press enter to stop): ")
        if name == "":
            break
        Net_ammount = int(input("   Ammount:    "))
        products = input("  Products:   ")
        quatity = int(input("   Quantity:   "))
        monthly = int(input("   Mothly Time:    "))
        Gpay = input("  Times:  ")
        History.append((name,Net_ammount,products,quatity,monthly,Gpay))
    return History

def create_workbook(History):

    wb = Workbook()
    ws = wb.active
    ws.title = "My Monthly Records"

    ws.append(["Name","Net_ammount","Products","Quantity","Monthly Time","Payment Method"])
    for row, package in enumerate(History, start=2):
        name, Net_ammount, products, quantity, monthly, Gpay = package
        ws.append([name,Net_ammount,products,quantity,monthly,Gpay])

    last_row = ws.max_row
    ws.append(["Grand Total",f"=SUM(B2:B{last_row})"])
    return wb


def main():
    History = to_ask_items()
    if not History:
        print("No items were added, nothing saved.")
        return
    wb = create_workbook(History)


    wb.save("Monthly_Dropout.xlsx")
    print("Records are saved")


main()