import pandas as pd
import json
import os

# Excel dosyasının yolu
excel_file = r'C:\Users\F\Desktop\.xlsx'

# Veri çekilecek sayfa
sheet_name = 'sayfa'

# JSON dosyalarının kaydedileceği dizin
output_dir = r'C:\Users\F\Desktop\Yeni klasör'

# Excel dosyasını oku
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# Boş bir sözlük oluştur
data_dict = {}

# Her satır için işlemleri yap
for index, row in df.iterrows():
    edas_maintenance_id = row['edas_maintenance_id']
    cost_items_name = row['cost_items_name']
    
    # Eğer edas_maintenance_id daha önce sözlüğe eklenmemişse, yeni bir girdi oluştur
    if edas_maintenance_id not in data_dict:
        data_dict[edas_maintenance_id] = {
            "edas_maintenance_id": edas_maintenance_id,
            "detail_cost": [],
            "maintenance_cost": 0.0,
            "labor_costs": 0.0
        }
    
    # Detail_cost içine yeni maliyet kalemi ekle
    detail_cost_item = {
        "cost_items_name": cost_items_name,
        "cost_type": row['cost_type'],
        "cost_items_value": row['cost_items_value'],
        "item_unit_value": row['item_unit_value'],
        "labor_costs": row['labor_costs'],
        "labor_costs_item_value": row['labor_costs_item_value'],
        "item_purchase_date": row['item_purchase_date'].strftime('%Y') if isinstance(row['item_purchase_date'], pd.Timestamp) else str(row['item_purchase_date']),
        "amount": row['amount']
    }
    
    data_dict[edas_maintenance_id]['detail_cost'].append(detail_cost_item)
    
    # Maintenance_cost ve labor_costs toplamlarını güncelle
    data_dict[edas_maintenance_id]['maintenance_cost'] += row['maintenance_cost']
    data_dict[edas_maintenance_id]['labor_costs'] += row['labor_costs']

# JSON dosyalarını oluştur
for edas_maintenance_id, data in data_dict.items():
    json_filename = os.path.join(output_dir, f"{edas_maintenance_id}.json")
    with open(json_filename, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=4)

print("JSON dosyaları oluşturuldu.")
