import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from xml.dom import minidom

# PNSYファイルの内容を文字列として生成する関数
def create_pnsy_content(sequences):
    root = ET.Element("DocY", xmlns_xsd="http://www.w3.org/2001/XMLSchema", xmlns_xsi="http://www.w3.org/2001/XMLSchema-instance")
    
    # <ListSequence>以前の要素を追加
    ET.SubElement(root, "Memo")
    ET.SubElement(root, "Iflg").text = "None"
    
    ipsel = ET.SubElement(root, "Ipsel")
    for _ in range(5):
        ET.SubElement(ipsel, "string")
    
    ET.SubElement(root, "Ds").text = "0.01"
    ET.SubElement(root, "NItu").text = "0"
    ET.SubElement(root, "Default").text = "true"
    ET.SubElement(root, "Ncmax").text = "999"
    ET.SubElement(root, "Dlt").text = "0.0001"
    ET.SubElement(root, "C").text = "Off"
    ET.SubElement(root, "F").text = "Normal"
    ET.SubElement(root, "Header")
    
    list_sequence = ET.SubElement(root, "ListSequence")
    
    for seq in sequences:
        sequence = ET.SubElement(list_sequence, "Sequence")
        
        ktyp = ET.SubElement(sequence, "Ktyp")
        ktyp.text = seq['Ktyp']

        phase = ET.SubElement(sequence, "Phase")
        if seq['Ktyp'] == 'N':
            phase.text = "N"
        elif seq['Ktyp'] == 'L':
            phase.text = "LCOE"
        elif seq['Ktyp'] == 'L2':
            phase.text = "LVAR"
        elif seq['Ktyp'] == 'S':
            phase.text = "GABC"
        elif seq['Ktyp'] == 'C':
            phase.text = "GABC"
        elif seq['Ktyp'] == 'G':
            phase.text = "GABC"
        elif seq['Ktyp'] == 'O':
            phase.text = "GABC"
        elif seq['Ktyp'] == 'A':
            phase.text = "GABC"
        
        time = ET.SubElement(sequence, "Time")
        time.text = str(seq['Time'])

        ichi = ET.SubElement(sequence, "Ichi")
        ichi.text = str(seq['Ichi'])
        
        name = ET.SubElement(sequence, "Name")
        sr = ET.SubElement(sequence, "Sr")
        if seq['Ktyp'] in ['S', 'G']:
            sr.text = "S"
        elif seq['Ktyp'] in ['C', 'O']:
            sr.text = "SR"
        
        ipos = ET.SubElement(sequence, "Ipos")
        
        ipsk = ET.SubElement(sequence, "Ipsk")
        ipsk.text = "10"
        
        ipskx = ET.SubElement(sequence, "Ipskx")
        ipskx.text = "1"
        
        if 'Data1' in seq and seq['Data1'] is not None:
            data1 = ET.SubElement(sequence, "Data1")
            data1.text = str(seq['Data1'])
        
        if 'Data2' in seq and seq['Data2'] is not None:
            data2 = ET.SubElement(sequence, "Data2")
            data2.text = str(seq['Data2'])
        
        if 'Data3' in seq and seq['Data3'] is not None:
            data3 = ET.SubElement(sequence, "Data3")
            data3.text = str(seq['Data3'])
        
        if 'Data4' in seq and seq['Data4'] is not None:
            data4 = ET.SubElement(sequence, "Data4")
            data4.text = str(seq['Data4'])

        if 'Data5' in seq and seq['Data5'] is not None:
            data5 = ET.SubElement(sequence, "Data5")
            data5.text = str(seq['Data5'])

        if 'Data6' in seq and seq['Data6'] is not None:
            data6 = ET.SubElement(sequence, "Data6")
            data6.text = str(seq['Data6'])
        
        data6num = ET.SubElement(sequence, "Data6Num")
    
    # ListSequence以下の要素を追加
    ET.SubElement(root, "Tmax").text = "20.0"
    ET.SubElement(root, "Agsm").text = "360.0"
    ET.SubElement(root, "Ngs1").text = "6"
    ET.SubElement(root, "Ngs2")
    ET.SubElement(root, "Ngs3")
    ET.SubElement(root, "Ngs4")
    ET.SubElement(root, "Ngs5")
    ET.SubElement(root, "Sgomax").text = "0.002"
    ET.SubElement(root, "IsFsb").text = "false"
    ET.SubElement(root, "Fsb2")
    ET.SubElement(root, "Fsb3")
    ET.SubElement(root, "Fsb4")
    ET.SubElement(root, "Fsb5")
    ET.SubElement(root, "DictDCexte")
    ET.SubElement(root, "DictDCself")
    ET.SubElement(root, "DictDCexte2")
    ET.SubElement(root, "DictDCself2")
    ET.SubElement(root, "ListDBlk")
    ET.SubElement(root, "ListDCst")
    ET.SubElement(root, "PowerControls")
    ET.SubElement(root, "Converters")
    ET.SubElement(root, "ListOgs")
    ET.SubElement(root, "PG").text = "2"
    
    list_pg = ET.SubElement(root, "ListPG")
    for i in range(1, 11):
        ET.SubElement(list_pg, "string").text = str(i)
    
    ET.SubElement(root, "GG").text = "2"
    
    list_gg = ET.SubElement(root, "ListGG")
    for i in range(1, 11):
        ET.SubElement(list_gg, "string").text = str(i)
    
    ET.SubElement(root, "PAVR").text = "1"
    ET.SubElement(root, "ListPA")
    ET.SubElement(root, "GAVR").text = "1"
    ET.SubElement(root, "ListGA")
    ET.SubElement(root, "PGOV").text = "1"
    ET.SubElement(root, "ListPV")
    ET.SubElement(root, "GGOV").text = "1"
    ET.SubElement(root, "ListGV")
    ET.SubElement(root, "PN").text = "2"
    
    list_pn = ET.SubElement(root, "ListPN")
    for i in range(1, 48):
        ET.SubElement(list_pn, "string").text = str(i)
    
    ET.SubElement(root, "GN").text = "2"
    
    list_gn = ET.SubElement(root, "ListGN")
    for i in range(1, 48):
        ET.SubElement(list_gn, "string").text = str(i)
    
    ET.SubElement(root, "PB").text = "2"
    
    list_pb = ET.SubElement(root, "ListPB")
    for i in list(range(1, 54)) + [117, 118, 120, 123, 124, 126, 135, 153]:
        ET.SubElement(list_pb, "string").text = str(i)
    
    ET.SubElement(root, "GB").text = "2"
    
    list_gb = ET.SubElement(root, "ListGB")
    for i in list(range(1, 54)) + [117, 118, 120, 123, 124, 126, 135, 153]:
        ET.SubElement(list_gb, "string").text = str(i)
    
    ET.SubElement(root, "PIM").text = "1"
    ET.SubElement(root, "ListPIM")
    ET.SubElement(root, "BOda").text = "false"
    ET.SubElement(root, "BOdea").text = "false"
    ET.SubElement(root, "BOsa").text = "false"
    ET.SubElement(root, "BOu").text = "false"
    ET.SubElement(root, "BOua").text = "false"
    ET.SubElement(root, "BOuas").text = "false"
    ET.SubElement(root, "BOuf").text = "false"
    ET.SubElement(root, "BOgd").text = "false"
    ET.SubElement(root, "Ogd")
    ET.SubElement(root, "ListOgd")
    ET.SubElement(root, "Fout").text = "10"
    ET.SubElement(root, "BZm").text = "false"
    ET.SubElement(root, "Zout").text = "10"
    ET.SubElement(root, "NSection").text = "3"
    
    yreport = ET.SubElement(root, "Yreport")
    ET.SubElement(yreport, "List")
    
    qdc = ET.SubElement(root, "Qdc")
    ET.SubElement(qdc, "List")
    
    ET.SubElement(root, "BRPrintCntl").text = "false"
    ET.SubElement(root, "PowerControlsOutPutDatasList")
    ET.SubElement(root, "BDcPrintCntl").text = "false"
    ET.SubElement(root, "DcOutPutDatasList")
    ET.SubElement(root, "BFr").text = "false"
    
    frs = ET.SubElement(root, "Frs")
    ET.SubElement(frs, "List")
    
    ET.SubElement(root, "BSor").text = "false"
    
    sors = ET.SubElement(root, "Sors")
    ET.SubElement(sors, "List")
    ET.SubElement(sors, "Pu").text = "true"
    
    ET.SubElement(root, "CardD")
    ET.SubElement(root, "CardGCHK")
    ET.SubElement(root, "CardGCON")
    ET.SubElement(root, "CardGSAT")
    ET.SubElement(root, "CardG")
    ET.SubElement(root, "CardA")
    ET.SubElement(root, "CardP")
    ET.SubElement(root, "CardS")
    ET.SubElement(root, "CardM")
    ET.SubElement(root, "CardR")
    ET.SubElement(root, "CardL")
    ET.SubElement(root, "CardF")
    ET.SubElement(root, "CardZ")
    
    docyc = ET.SubElement(root, "Docyc")
    ET.SubElement(docyc, "Name")
    ET.SubElement(docyc, "List")
    ET.SubElement(docyc, "ListL")
    
    # 出力を整形
    rough_string = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")

# 各行を個別のシーケンスとして処理し、すべてのPNSYファイルを1つのPOPファイルにまとめる関数
def process_excel_file(file_path, start_index=0):
    xlsx_data = pd.read_excel(file_path, sheet_name=None, dtype=str)
    current_directory = os.getcwd()
    pop_file_name = os.path.basename(file_path).replace('.xlsx', '.pop')
    pop_file_path = os.path.join(current_directory, pop_file_name)
    
    with zipfile.ZipFile(pop_file_path, 'w') as zipf:
        for sheet_name, sheet_data in xlsx_data.items():
            s = 0
            for row_index, row in sheet_data.iterrows():
                s = start_index + row_index

                # Time列が文字列であることを確認
                if isinstance(row['Time'], str):
                    time_list = row['Time'].split(',')
                else:
                    # 数値の場合、エラーを避けるために適切に処理
                    print(f"Warning: 'Time' at row {row_index} is not a string. Value: {row['Time']}")
                    continue
                
                ktyp_list = row['Ktyp'].split(',')
                ichi_list = row['Ichi'].split(',')
                data1_list = row['Data1'].split(',') if isinstance(row['Data1'], str) else ([row['Data1']] if pd.notna(row['Data1']) else [])
                data2_list = row['Data2'].split(',') if isinstance(row['Data2'], str) else ([row['Data2']] if pd.notna(row['Data2']) else [])
                data3_list = row['Data3'].split(',') if isinstance(row['Data3'], str) else ([row['Data3']] if pd.notna(row['Data3']) else [])
                data4_list = row['Data4'].split(',') if isinstance(row['Data4'], str) else ([row['Data4']] if pd.notna(row['Data4']) else [])
                if 'Data5' in row:
                    data5_list = row['Data5'].split(',') if isinstance(row['Data5'], str) else ([row['Data5']] if pd.notna(row['Data5']) else [])
                if 'Data6' in row:
                    data6_list = row['Data6'].split(',') if isinstance(row['Data6'], str) else ([row['Data6']] if pd.notna(row['Data6']) else [])
                
                sequences = []
                for i in range(len(time_list)):
                    if 'Data5' in row:
                        sequence = {
                            'Time': time_list[i],
                            'Ktyp': ktyp_list[i] if i < len(ktyp_list) else '',
                            'Ichi': ichi_list[i] if i < len(ichi_list) else '',
                            'Data1': data1_list[i] if i < len(data1_list) else None,
                            'Data2': data2_list[i] if i < len(data2_list) else None,
                            'Data3': data3_list[i] if i < len(data3_list) else None,
                            'Data4': data4_list[i] if i < len(data4_list) else None,
                            'Data5': data5_list[i] if i < len(data5_list) else None,
                            'Data6': data6_list[i] if i < len(data6_list) else None
                        }
                    else:
                        sequence = {
                            'Time': time_list[i],
                            'Ktyp': ktyp_list[i] if i < len(ktyp_list) else '',
                            'Ichi': ichi_list[i] if i < len(ichi_list) else '',
                            'Data1': data1_list[i] if i < len(data1_list) else None,
                            'Data2': data2_list[i] if i < len(data2_list) else None,
                            'Data3': data3_list[i] if i < len(data3_list) else None,
                            'Data4': data4_list[i] if i < len(data4_list) else None,
                        }
                    
                    sequences.append(sequence)
                
                pnsy_content = create_pnsy_content(sequences)
                
                if s < 10:
                    pnsy_filename = f'0{s}_{"_".join(ichi_list)}.pnsy'
                else:
                    pnsy_filename = f'{s}_{"_".join(ichi_list)}.pnsy'

                zipf.writestr(pnsy_filename, pnsy_content)
    
    return pop_file_path


# ファイルパスを手入力するためのコードを追加
if __name__ == "__main__":
    file_path = 'condition0.xlsx'
    start_index = 1
    pop_file_path = process_excel_file(file_path, start_index)
    print(f"解析対象条件が記載されたPOPファイル: {pop_file_path}")
