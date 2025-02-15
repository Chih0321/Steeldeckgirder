import pandas as pd
import re
import math

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

def Girdersection(secid, df_section):
    # 計算頂板
    B_top = df_section['B1'][secid] +df_section['B2'][secid] +df_section['B3'][secid] 
    t1 = df_section['t1'][secid]

    area_top = B_top*t1
    iyy_top = B_top*(t1**3)/12
    izz_top = t1*(B_top**3)/12

    z2top_top = t1/2
    y2ref_top = df_section['Ref_top'][secid] +B_top/2

    # 計算左腹板
    H = df_section['H'][secid]
    tw1 = df_section['tw1'][secid]

    difference_y1 = df_section['Ref_top'][secid] +df_section['B1'][secid] -df_section['Ref_bot'][secid] -df_section['B4'][secid]
    inclineangle1 = math.atan(difference_y1 /H)
    thicknesswide1 = tw1/math.cos(inclineangle1)

    area_web1 = thicknesswide1*H
    iyy_web1 = thicknesswide1*(H**3)/12
    izz_web1 = H*(thicknesswide1**3)/12

    z2top_web1 = t1 +H/2
    y2ref_web1 = (df_section['Ref_top'][secid] +df_section['B1'][secid] +df_section['Ref_bot'][secid] +df_section['B4'][secid])/2 -thicknesswide1/2

    # 計算右腹板
    tw2 = df_section['tw2'][secid]

    difference_y2 = df_section['Ref_top'][secid] +df_section['B1'][secid] +df_section['B2'][secid] -df_section['Ref_bot'][secid] -df_section['B4'][secid] -df_section['B5'][secid]
    inclineangle2 = math.atan(difference_y2 /H)
    thicknesswide2 = tw2/math.cos(inclineangle2)

    area_web2 = thicknesswide2*H
    iyy_web2 = thicknesswide2*(H**3)/12
    izz_web2 = H*(thicknesswide2**3)/12

    z2top_web2 = t1 +H/2
    y2ref_web2 = (df_section['Ref_top'][secid] +df_section['B1'][secid] +df_section['B2'][secid] +df_section['Ref_bot'][secid] +df_section['B4'][secid] +df_section['B5'][secid])/2 +thicknesswide2/2

    # 計算底版
    B_bot = df_section['B4'][secid] +df_section['B5'][secid] +df_section['B6'][secid]
    t2 = df_section['t2'][secid]

    area_bot = B_bot*t2
    iyy_bot = B_bot*(t2**3)/12
    izz_bot = t2*(B_bot**3)/12

    z2top_bot = t1 +H +t2/2
    y2ref_bot = df_section['Ref_bot'][secid] +B_bot/2

    # 總斷面
    ## 面積
    area_all = area_top +area_web1 +area_web2 +area_bot
    ## 中性軸
    z_na = (area_top*z2top_top +area_web1*z2top_web1 +area_web2*z2top_web2 +area_bot*z2top_bot) /area_all
    y_na = (area_top*y2ref_top +area_web1*y2ref_web1 +area_web2*y2ref_web2 +area_bot*y2ref_bot) /area_all
    ## 慣性矩
    iyy_all = iyy_top +area_top*(z_na -z2top_top)**2 +iyy_web1 +area_web1*(z_na -z2top_web1)**2 +iyy_web2 +area_web2*(z_na -z2top_web2)**2 +iyy_bot +area_bot*(z_na -z2top_bot)**2
    izz_all = izz_top +area_top*(y_na -y2ref_top)**2 +izz_web1 +area_web1*(y_na -y2ref_web1)**2 +izz_web2 +area_web2*(y_na -y2ref_web2)**2 +izz_bot +area_bot*(y_na -y2ref_bot)**2

    return area_all, z_na, y_na, iyy_all, izz_all



# %% 讀取輸入
inputfile = r"/Users/chih/Documents/Code/Steeldeckgirder/Section.xlsx"
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

# %% 計算主梁斷面
# for run_id in range(len(df_section)):
#     area_girder, z_girder, y_girder, iyy_girder, izz_girder = Girdersection(run_id, df_section)

# %% 計算加勁鈑
rib_id = 0
r_id = rib_id

if df_ribs['Type'][r_id] == 'Flat':
    area_r = df_ribs['H/H/H'][r_id]*df_ribs['B/B/B1'][r_id]
    z_r = df_ribs['H/H/H'][r_id]/2
    y_r = 0
    iyy_r = df_ribs['B/B/B1'][r_id]*(df_ribs['H/H/H'][r_id])**3/12
    izz_r = df_ribs['H/H/H'][r_id]*(df_ribs['B/B/B1'][r_id])**3/12
elif df_ribs['Type'][r_id] == 'Tee':
    bt = df_ribs['B/B/B1'][r_id]
    tft = df_ribs['/tf/t'][r_id]
    ht = df_ribs['H/H/H'][r_id]
    twt = df_ribs['/tw/B2'][r_id]
    area_r = bt*tft +(ht-tft)*twt

    z_r = (ht-tft)*twt*(ht-tft)/2 +bt*tft*(ht-tft/2) 
    y_r = 0

    iyy_r = twt*(ht-tft)**3/12 +twt*(ht-tft)*(z_r-(ht-tft)/2)**2 +bt*tft**3/12 +bt*tft*(z_r-(ht-tft/2) )**2
    izz_r = (ht-tft)*twt**3/12 +tft*bt**3/12
print('break point')

