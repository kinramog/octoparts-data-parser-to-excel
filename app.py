import os, csv
from nexarClient import NexarClient
import xlsxwriter
import PySimpleGUI as sg

def write_to_excel(results, local_batch):
    if results:
        wb = xlsxwriter.Workbook('parts.xlsx')
        ws = wb.add_worksheet('Pricing breaks')
        
        mpn_format = wb.add_format({'bold': True, 'font_size': 14})
        seller_format = wb.add_format({'bold': True, 'border': 1})
        general_format = wb.add_format({'border': 1})
        seller_format_chosen = wb.add_format({'bold': True, 'border': 1, "bg_color": "#D9D9D9"})
        general_format_chosen = wb.add_format({'border': 1, "bg_color": "#D9D9D9"})
        start_MPN, end_MPN = 0, 0
        max_width = 20
        for id, res in enumerate(results):
            query = res.get('results')[0]

            ws.write(start_MPN, end_MPN, 'MPN', mpn_format)
            ws.write(start_MPN, end_MPN+1, 'Name', mpn_format)
            ws.write(start_MPN+1, end_MPN, query.get("part").get("mpn"), mpn_format)
            ws.write(start_MPN+1, end_MPN+1, query.get("part").get("name"), mpn_format)

            if len(query.get("part").get("mpn")) > max_width:
                max_width = len(query.get("part").get("mpn"))
            
            row = 3 if id == 0 else start_MPN+3
            for seller in query.get("part").get("sellers"):
                ws.write(row, 0, "Seller name", seller_format)
                ws.write(row+1, 0, seller.get("company").get("name"), general_format)

                for offer in seller.get("offers"):
                    ws.write(row, 1, "Stock", seller_format)
                    ws.write(row+1, 1, offer.get("inventoryLevel"), general_format)
                    # for i, price in enumerate(offer.get("prices")):
                    #     ws.write(row, 1 + (i + 1), price.get("quantity"), seller_format)
                    #     ws.write(row+1, 1 + (i + 1), price.get("price"), general_format)
                        
                    prices = offer.get("prices")
                    for i in range(len(prices)-2):
                        local_batch_value = local_batch.get(id) if local_batch.get(id) else -1000
                        
                        if prices[i].get("quantity") < local_batch_value < prices[i+1].get("quantity"):
                            ws.write(row, 1 + (i + 1), prices[i].get("quantity"), seller_format_chosen)
                            ws.write(row+1, 1 + (i + 1), prices[i].get("price"), general_format_chosen)
                        else:
                            ws.write(row, 1 + (i + 1), prices[i].get("quantity"), seller_format)
                            ws.write(row+1, 1 + (i + 1), prices[i].get("price"), general_format)
                row +=3                 
            
            start_MPN = row
        ws.set_column("A:A", max_width)
        wb.close()
    else:
        print("errorreee, no reesults")


gqlQuery = '''
query pricingByVolumeLevels($mpn: String!) {
  supSearchMpn(q: $mpn, limit: 1) {
    results {
      part {
        mpn
        name
        sellers {
          company {
            name
          }
          offers {
            inventoryLevel
            prices {
              quantity
              price
            }
          }
        }
      }
    }
  }
}
'''

layout = [
    [sg.Text("Вставьте партномера: ")],
    [
        [sg.Multiline(size=(30, 10), expand_x=True, expand_y=True),
        sg.Button("Получить данные")], 
        [sg.Text('', key='-OUTPUT-')] 
    ],
]

sg.theme("Gray Gray Gray")
window = sg.Window("Parts parser", layout, resizable=True, element_justification="l", finalize=True)
window.set_min_size(size=(400,200))
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        sg.PopupAnimated(None)
        break

    if values[0] != "":
        sg.PopupAnimated(sg.DEFAULT_BASE64_LOADING_GIF, background_color='white', time_between_frames=100)

        raw_mpns = list(filter(lambda x: x != "", values[0].split("\n")))
        print(raw_mpns)
        local_batch = {}
        mpns = []
        for id, mpn in enumerate(raw_mpns):
            if len(mpn.split()) > 1:
                local_batch[id] = int(mpn.split()[1])
                mpns.append(mpn.split()[0])
            else:
                mpns.append(mpn)

        clientId = ""
        clientSecret = ""
        with open("config.txt", "r") as file:
            config = file.readlines()
            clientId = config[0][len("clientID="):].strip()
            clientSecret = config[1][len("clientSecret="):].strip()
        
        # clientId = "8d6d66c8-016e-4cef-830c-d47b9babd536"
        # clientSecret = "Tk-XOx2MeJs2OJgjtCSmYu_QrZUx-moLaIeO"
        nexar = NexarClient(clientId, clientSecret)

        results = []
        for mpn in mpns:
            try:
                variables = {"mpn": mpn}
                result = nexar.get_query(gqlQuery, variables)
                results.append(result.get('supSearchMpn'))
            except Exception as e:
                print("ERROR in result = nexar.get_query(gqlQuery, variables)\n", e)
        
        write_to_excel(results, local_batch)
        sg.PopupAnimated(None)
        window['-OUTPUT-'].update("Данные получены и записаны в файл parts.xlsx")
		

window.close()