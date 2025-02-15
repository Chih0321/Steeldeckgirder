import pandas as pd
import re

def Generatemct(df_sec, df_rs, num_id):
    """生成Steel Girder MCT指令

    Args:
        df_sec (DataFrame): 輸入之鋼梁斷面資訊
        df_rs (DataFrame): 加勁鈑輸入之斷面庫

    Returns:
        str: 輸出MCT用指令
    """
    section_id = str(df_sec['編號'][num_id])
    section_type = 'SOD'
    section_name = str(df_sec['Name'][num_id])
    section_offset = 'CC'
    section_shape = 'SOD-B'

    section_B1 = str(df_sec['B1'][num_id])
    section_B2 = str(df_sec['B2'][num_id])
    section_B3 = str(df_sec['B3'][num_id])
    section_B4 = str(df_sec['B4'][num_id])
    section_B5 = str(df_sec['B5'][num_id])
    section_B6 = str(df_sec['B6'][num_id])
    section_H = str(df_sec['H'][num_id])
    section_t1 = str(df_sec['t1'][num_id])
    section_t2 = str(df_sec['t2'][num_id])
    section_tw1 = str(df_sec['tw1'][num_id])
    section_tw2 = str(df_sec['tw2'][num_id])
    section_reftop = str(df_sec['Ref_top'][num_id])
    section_refbot = str(df_sec['Ref_bot'][num_id])

    section_numribs = str(df_rs.shape[0])
    rib_data = ''
    for r, r_info in df_rs.iterrows():
        r_name = r_info['Name']
        r_type = '0' if r_info['Type'] == 'Flat' else '1' if r_info['Type'] == 'Tee' else '2'
        r_section = ', '.join([str(r_info['H/H/H']), str(r_info['B/B/B1']), str(r_info['/tw/B2']), str(r_info['/tf/t']), str(r_info['//R']), '0, 0, 0'])

        rib_data = rib_data +', ' +', '.join([r_name, r_type, r_section])

    def command_ribs_position(rdst, rdp, rds, rdd, rdn, rdpos, nid):

        rib_type = df_sec[rdp][nid]
        rib_spacing = re.split(r'\s*,\s*', df_sec[rds][nid])
        rib_num = len(rib_spacing)

        rib_pos = ''    
        for rr in range(len(rib_spacing)):
            command_ribpos = ', '.join(['YES', rib_spacing[rr], rib_type, rdpos, rdn+str(rr+1)])
            rib_pos = rib_pos +', ' +command_ribpos

        command = rdd +', ' +rdst +', 0, ' +str(rib_num) +', ' +str(rib_num) +rib_pos

        return command

    rib_define_situtaion = ['Top-Left', 'Top-Center', 'Top-Right', 'Bottom-Left', 'Bottom-Center', 'Bottom-Right']
    rib_define_position = ['Top-Flange (Left-type)', 'Top-Flange (Center-type)', 'Top-Flange (Right-type)', 'Bottom-Flange (Left-type)', 'Bottom-Flange (Center-type)', 'Bottom-Flange (Right-type)']
    rib_define_spacing = ['Top-Flange (Left-spacing)', 'Top-Flange (Center-spacing)', 'Top-Flange (Right-spacing)', 'Bottom-Flange (Left-spacing)', 'Bottom-Flange (Center-spacing)', 'Bottom-Flange (Right-spacing)']
    rib_define_deckid = ['0,0', '0,1', '0,2', '3,0', '3,1', '3,2']
    rib_define_name = ['TL', 'TC', 'TR', 'BL', 'BC', 'BR']
    rib_define_pos = ['1', '1', '1', '0', '0', '0']
    ribcommand = ''
    rib_poscount = 0
    for i in range(len(rib_define_situtaion)):
        if not pd.isna(df_sec[rib_define_position[i]][num_id]):
            rib_poscount = rib_poscount +1
            command_return = command_ribs_position(rib_define_situtaion[i], rib_define_position[i], rib_define_spacing[i], rib_define_deckid[i], rib_define_name[i], rib_define_pos[i], num_id)
            ribcommand = ribcommand + ', ' + command_return


    # ID, TYPE, Name, Offset, iCENT, iREF, iHORZ, HUSER, iVERT, VUSER, bSD, bWE, SHAPE  
    line_common = ', '.join([section_id, section_type, section_name, section_offset, '0, 0, 0, 0, 0, 0, YES, NO', section_shape])
    # AUTOSYM, B1, B2, B3, B4, B5, B6, H, t1, t2,  tw1, tw2, DRLTop, DRLBot
    line_dimension = ', '.join(['NO', section_B1, section_B2, section_B3, section_B4, section_B5, section_B6, section_H, section_t1, section_t2, section_tw1, section_tw2, section_reftop, section_refbot ])
    # numRibDB, name, TYPE, H, B, tw, tf
    line_ribsec = section_numribs + rib_data
    # numPos, 0, 0, POSITION, LEFT(0)/RIGHT(1), numRib, numRib, C, SPACE, Rid, position, Name 
    line_ribpos = str(rib_poscount) +ribcommand

    line_all = line_common +'\n   ' +line_dimension +'\n   ' +line_ribsec +'\n   ' +line_ribpos

    return line_all




# %% 讀取輸入
inputfile = r"D:\Users\63427\Desktop\Code\鋼床鈑\Steeldeckgirder\Section.xlsx"
df_section = pd.read_excel(inputfile, sheet_name='鋼床鈑', skiprows=[1])
df_ribs = pd.read_excel(inputfile, sheet_name='加勁鈑')


# %% 寫MCT指令
# commandmct = ""
# for run_id in range(len(df_section)):

#     commandmct_single = Generatemct(df_section, df_ribs, run_id)
#     if commandmct == "":
#         commandmct = commandmct_single
#     else:
#         commandmct = commandmct +'\n' +commandmct_single

# mctcommandfile = "MCT_STLB.txt"
# with open(mctcommandfile, "w", encoding="utf-8") as file:
#     file.write(commandmct)
# print(f"$ 字串已成功寫入 {mctcommandfile} 文件。")

# %% 自行計算斷面
run_id =0

secid = run_id

B_top = df_section['B1'][secid] +df_section['B2'][secid] +df_section['B3'][secid] 
B_bot = df_section['B4'][secid] +df_section['B5'][secid] +df_section['B6'][secid] 
print('break point')

