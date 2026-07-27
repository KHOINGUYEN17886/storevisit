import json

with open(r'c:\All_Report\8_RETAIL_COMMANDER\StoreVisit\webapp\store_mapping.json', 'r', encoding='utf-8') as f:
    d = json.load(f)

# Build the JS literal for STORE_DATA_MAP
lines = []
lines.append('var STORE_DATA_MAP = {')
lines.append('  "mapping_by_asm": {')
asm_items = list(d['mapping_by_asm'].items())
for idx, (asm, stores) in enumerate(asm_items):
    comma = ',' if idx < len(asm_items)-1 else ''
    stores_str = json.dumps(stores, ensure_ascii=False)
    lines.append('      ' + json.dumps(asm, ensure_ascii=False) + ': ' + stores_str + comma)
lines.append('  },')
lines.append('  "mapping_by_region": {')
reg_items = list(d['mapping_by_region'].items())
for idx, (reg, stores) in enumerate(reg_items):
    comma = ',' if idx < len(reg_items)-1 else ''
    stores_str = json.dumps(stores, ensure_ascii=False)
    lines.append('      ' + json.dumps(reg, ensure_ascii=False) + ': ' + stores_str + comma)
lines.append('  }')
lines.append('};')

output = '\n'.join(lines)

with open(r'c:\All_Report\8_RETAIL_COMMANDER\StoreVisit\webapp\store_data_map_new.js', 'w', encoding='utf-8') as f:
    f.write(output)
print('Saved store_data_map_new.js, lines:', len(lines))
print('ASMs:', list(d['mapping_by_asm'].keys()))
print('Regions:', list(d['mapping_by_region'].keys()))
