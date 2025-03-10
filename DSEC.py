import pandas as pd
import numpy as np
import re
import math
import os
import ast
import ezdxf
from ezdxf.enums import TextEntityAlignment


import sys
from PySide6 import QtCore
from PySide6.QtCore import QFile, QThread, QObject, Signal, QEventLoop, QTimer
from PySide6.QtUiTools import QUiLoader
from PySide6.QtGui import QIcon, QTextCursor
from PySide6.QtWidgets import QApplication, QFileDialog
import ctypes
myappid = 'SteelDeck Secion Calculator' # arbitrary string
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


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
    section_offset = 'CT' # or CC
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


def Mctmainexe(inputfile):
    """執行MCT生成主要DEF

    Args:
        inputfile (str): 輸入參數的excel檔
    """
    # %% 寫MCT指令
    print("$ 執行生成斷面(STLB)MCT指令。")
    # 讀檔
    (outputpath, filename_temp) = os.path.split(inputfile)
    inputfilename = filename_temp.split(".")[0]

    df_section = pd.read_excel(inputfile, sheet_name='鋼床鈑', skiprows=[1])
    df_ribs = pd.read_excel(inputfile, sheet_name='加勁鈑')

    # 生成MCT
    commandmct = ""
    for run_id in range(len(df_section)):

        commandmct_single = Generatemct(df_section, df_ribs, run_id)
        if commandmct == "":
            commandmct = commandmct_single
        else:
            commandmct = commandmct +'\n' +commandmct_single

    mctcommandfile = os.path.join(outputpath,inputfilename+"_MCT_STLB.txt")
    with open(mctcommandfile, "w", encoding="utf-8") as file:
        file.write(commandmct)

    print("> MCT指令已成功寫入 {}。".format(inputfilename+"_MCT_STLB.txt"))


def Girdersection(B_top, t1, B_bot, t2, H, tw1_i, tw2_i, ref_top, ref_bot, B1, B2, B4, B5):
    """主梁斷面計算(不含加勁)

    Args:
        B_top (float): 頂版寬/有效頂版寬
        t1 (float): 頂版厚
        B_bot (float): 底版寬/有效底版寬
        t2 (float): 底版厚
        H (float): 淨高
        tw1_i (float): 左腹鈑厚
        tw2_i (float): 右腹鈑厚
        ref_top (float): 頂版位置參考距離
        ref_bot (float): 底版位置參考距離
        B1 (float): 頂版B1
        B2 (float): 頂版B2
        B4 (float): 底版B4
        B5 (float): 底版B5

    Returns:
        float: 計算之斷面積, 中性軸位置, 慣性矩
    """

    # 計算頂板
    area_top = B_top*t1
    iyy_top = B_top*(t1**3)/12
    izz_top = t1*(B_top**3)/12

    z2top_top = t1/2
    y2ref_top = ref_top +B_top/2

    # 計算左腹板
    area_web1 = tw1_i*H
    iyy_web1 = tw1_i*(H**3)/12
    izz_web1 = H*(tw1_i**3)/12

    z2top_web1 = t1 +H/2
    y2ref_web1 = (ref_top +B1 +ref_bot +B4)/2 -tw1_i/2

    # 計算右腹板
    area_web2 = tw2_i*H
    iyy_web2 = tw2_i*(H**3)/12
    izz_web2 = H*(tw2_i**3)/12

    z2top_web2 = t1 +H/2
    y2ref_web2 = (ref_top +B1 +B2 +ref_bot +B4 +B5)/2 +tw2_i/2

    # 計算底版
    area_bot = B_bot*t2
    iyy_bot = B_bot*(t2**3)/12
    izz_bot = t2*(B_bot**3)/12

    z2top_bot = t1 +H +t2/2
    y2ref_bot = ref_bot +B_bot/2

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


def Ribproperty(df_ribs):
    """計算RIB斷面性質

    Args:
        df_ribs (DataFrame): 輸入加勁鈑斷面資料庫

    Returns:
        dict: 加勁鈑斷面性質
    """
    dict_r_property = {'Name':[],
                    'area':[],
                    'z_p_na':[],
                    'z_n_na':[],
                    'y_na':[],
                    'iyy':[],
                    'izz':[],
                    }
    for r_id in range(len(df_ribs)):
        if df_ribs['Type'][r_id] == 'Flat':
            area_r = df_ribs['H/H/H'][r_id]*df_ribs['B/B/B1'][r_id]
            z_p_r = df_ribs['H/H/H'][r_id]/2
            z_n_r = df_ribs['H/H/H'][r_id]/2
            y_r = df_ribs['B/B/B1'][r_id]/2

            iyy_r = df_ribs['B/B/B1'][r_id]*(df_ribs['H/H/H'][r_id])**3/12
            izz_r = df_ribs['H/H/H'][r_id]*(df_ribs['B/B/B1'][r_id])**3/12

            dict_r_property['Name'].append(df_ribs['Name'][r_id])
            dict_r_property['area'].append(area_r)
            dict_r_property['z_p_na'].append(z_p_r)
            dict_r_property['z_n_na'].append(z_n_r)
            dict_r_property['y_na'].append(y_r)
            dict_r_property['iyy'].append(iyy_r)
            dict_r_property['izz'].append(izz_r)

        elif df_ribs['Type'][r_id] == 'Tee':
            bt = df_ribs['B/B/B1'][r_id]
            tft = df_ribs['/tf/t'][r_id]
            ht = df_ribs['H/H/H'][r_id]
            twt = df_ribs['/tw/B2'][r_id]
            area_r = bt*tft +(ht-tft)*twt

            z_p_r = ((ht-tft)*twt*((ht-tft)/2+tft) +bt*tft*(tft/2))/area_r
            z_n_r = ht-z_p_r
            y_r = bt/2

            iyy_r = twt*(ht-tft)**3/12 +twt*(ht-tft)*(z_p_r-((ht-tft)/2+tft))**2 +bt*tft**3/12 +bt*tft*(z_p_r-(tft/2) )**2
            izz_r = (ht-tft)*twt**3/12 +tft*bt**3/12

            dict_r_property['Name'].append(df_ribs['Name'][r_id])
            dict_r_property['area'].append(area_r)
            dict_r_property['z_p_na'].append(z_p_r)
            dict_r_property['z_n_na'].append(z_n_r)
            dict_r_property['y_na'].append(y_r)
            dict_r_property['iyy'].append(iyy_r)
            dict_r_property['izz'].append(izz_r)

    return dict_r_property


def Includeribs(df_section, df_rib_property, run_id, section_area, section_zna, section_yna, section_iyy, section_izz, dict_position_top, dict_position_bot):
    ## 頂版加勁
    position_rib_top_y = []
    area_rib_top = []
    for key, item in dict_position_top.items():
        if not pd.isna(df_section[key][run_id]):
            ### 單獨rib斷面性質提取
            rib_type = df_section[key][run_id]
            rib_area = df_rib_property['area'][rib_type]
            rib_z_p = df_rib_property['z_p_na'][rib_type]
            rib_z_n = df_rib_property['z_n_na'][rib_type]
            rib_y = df_rib_property['y_na'][rib_type]
            rib_iyy = df_rib_property['iyy'][rib_type]
            rib_izz = df_rib_property['izz'][rib_type]

            ### 提取ribs間距
            rib_spacing = re.split(r'\s*,\s*', df_section[item[0]][run_id])
            rib_num = len(rib_spacing)

            ### 參考位置定位與初始化
            rib_z_level = item[1]
            rib_y_level = item[2]
            rib_dist_y = 0
            rib_area_top_temp = 0
            ### 合成性質計算
            for rr in range(len(rib_spacing)):
                #### rib座標位置定出
                set_z = rib_z_level + rib_z_n
                rib_dist_y = rib_dist_y + float(rib_spacing[rr])
                set_y = rib_y_level + rib_dist_y
                #### 計算考慮範圍內rib
                if not (item[3][0] < set_y < item[3][1]):
                    area_pre = section_area
                    zna_pre = section_zna
                    yna_pre = section_yna

                    section_area = section_area + rib_area
                    rib_area_top_temp = rib_area_top_temp + rib_area
                    area_rib_top.append(rib_area)                

                    position_rib_top_z = set_z
                    position_rib_top_y.append(set_y)

                    section_zna = ((area_pre)*section_zna + rib_area*set_z)/section_area
                    section_yna = ((area_pre)*section_yna + rib_area*set_y)/section_area

                    section_iyy = section_iyy + area_pre*(section_zna - zna_pre)**2 + rib_iyy + rib_area*(section_zna - set_z)**2
                    section_izz = section_izz + area_pre*(section_yna - yna_pre)**2 + rib_izz + rib_area*(section_yna - set_y)**2

    ## 底版加勁
    position_rib_bot_y = []
    area_rib_bot = []
    for key, item in dict_position_bot.items():
        if not pd.isna(df_section[key][run_id]):
            ### 單獨rib斷面性質提取
            rib_type = df_section[key][run_id]
            rib_area = df_rib_property['area'][rib_type]
            rib_z_p = df_rib_property['z_p_na'][rib_type]
            rib_z_n = df_rib_property['z_n_na'][rib_type]
            rib_y = df_rib_property['y_na'][rib_type]
            rib_iyy = df_rib_property['iyy'][rib_type]
            rib_izz = df_rib_property['izz'][rib_type]

            ### 提取ribs間距
            rib_spacing = re.split(r'\s*,\s*', df_section[item[0]][run_id])
            rib_num = len(rib_spacing)

            ### 參考位置定位與初始化
            rib_z_level = item[1]
            rib_y_level = item[2]
            rib_dist_y = 0
            rib_area_bot_temp = 0
            ### 合成性質計算
            for rr in range(len(rib_spacing)):
                #### rib座標位置定出
                set_z = rib_z_level - rib_z_n
                rib_dist_y = rib_dist_y + float(rib_spacing[rr])
                set_y = rib_y_level + rib_dist_y
                #### 計算考慮範圍內rib
                if not (item[3][0] < set_y < item[3][1]):
                    area_pre = section_area
                    zna_pre = section_zna
                    yna_pre = section_yna

                    section_area = section_area + rib_area
                    rib_area_bot_temp = rib_area_bot_temp + rib_area
                    area_rib_bot.append(rib_area)

                    position_rib_bot_z = set_z
                    position_rib_bot_y.append(set_y)

                    section_zna = ((area_pre)*section_zna + rib_area*set_z)/section_area
                    section_yna = ((area_pre)*section_yna + rib_area*set_y)/section_area

                    section_iyy = section_iyy + area_pre*(section_zna - zna_pre)**2 + rib_iyy + rib_area*(section_zna - set_z)**2
                    section_izz = section_izz + area_pre*(section_yna - yna_pre)**2 + rib_izz + rib_area*(section_yna - set_y)**2                

    rib_info_top = [rib_area_top_temp, area_rib_top, position_rib_top_z, position_rib_top_y]
    rib_info_bot = [rib_area_bot_temp, area_rib_bot, position_rib_bot_z, position_rib_bot_y]

    return section_area, section_zna, section_yna, section_iyy, section_izz, rib_info_top, rib_info_bot


def Sectioncalculation(inputfile):
    """自行計算鋼床鈑斷面

    Args:
        inputfile (str): 輸入參數的excel
    """
    print("$ 執行斷面(STLB)計算。")
    # %% 讀檔
    (outputpath, filename_temp) = os.path.split(inputfile)
    inputfilename = filename_temp.split(".")[0]

    df_section = pd.read_excel(inputfile, sheet_name='鋼床鈑', skiprows=[1])
    df_ribs = pd.read_excel(inputfile, sheet_name='加勁鈑')

    # %% 計算加勁鈑
    dict_rib_property = Ribproperty(df_ribs)
    df_rib_property = pd.DataFrame.from_dict(dict_rib_property)
    df_rib_property = df_rib_property.set_index('Name')    

    # %% 計算主梁斷面
    dict_section = {'Name':[],
                    'Area':[],
                    'Asy':[],
                    'Asz':[],
                    'Ixx':[],
                    'Iyy':[],
                    'Izz':[],
                    'yna_right':[],
                    'yna_left':[],
                    'zna_top':[],
                    'zna_bot':[],
                    'Zyy':[],
                    'Zzz':[],
                    }
    for run_id in range(len(df_section)):
        '''主要箱梁斷面性質'''
        H = df_section['H'][run_id]
        tw1 = df_section['tw1'][run_id]
        tw2 = df_section['tw2'][run_id]
        difference_y1 = df_section['Ref_top'][run_id] +df_section['B1'][run_id] -df_section['Ref_bot'][run_id] -df_section['B4'][run_id]
        inclineangle1 = math.atan(difference_y1 /H)
        thicknesswide1 = tw1/math.cos(inclineangle1)
        difference_y2 = df_section['Ref_top'][run_id] +df_section['B1'][run_id] +df_section['B2'][run_id] -df_section['Ref_bot'][run_id] -df_section['B4'][run_id] -df_section['B5'][run_id]
        inclineangle2 = math.atan(difference_y2 /H)
        thicknesswide2 = tw2/math.cos(inclineangle2)

        area_girder, z_girder, y_girder, iyy_girder, izz_girder = Girdersection(df_section['B1'][run_id] +df_section['B2'][run_id] +df_section['B3'][run_id], 
                                                                                df_section['t1'][run_id], 
                                                                                df_section['B4'][run_id] +df_section['B5'][run_id] +df_section['B6'][run_id],
                                                                                df_section['t2'][run_id], 
                                                                                df_section['H'][run_id], 
                                                                                thicknesswide1, 
                                                                                thicknesswide2, 
                                                                                df_section['Ref_top'][run_id], 
                                                                                df_section['Ref_bot'][run_id], 
                                                                                df_section['B1'][run_id], 
                                                                                df_section['B2'][run_id], 
                                                                                df_section['B4'][run_id], 
                                                                                df_section['B5'][run_id]
                                                                                )

        '''加入加勁鈑'''
        dict_position_top = {'Top-Flange (Left-type)':['Top-Flange (Left-spacing)', df_section['t1'][run_id], df_section['Ref_top'][run_id], [df_section['Ref_top'][run_id], df_section['Ref_top'][run_id]]],
                            'Top-Flange (Center-type)':['Top-Flange (Center-spacing)', df_section['t1'][run_id], df_section['Ref_top'][run_id]+df_section['B1'][run_id], [df_section['Ref_top'][run_id], df_section['Ref_top'][run_id]]],
                            'Top-Flange (Right-type)':['Top-Flange (Right-spacing)', df_section['t1'][run_id], df_section['Ref_top'][run_id]+df_section['B1'][run_id]+df_section['B2'][run_id], [df_section['Ref_top'][run_id], df_section['Ref_top'][run_id]]],
                            }
        dict_position_bot = {'Bottom-Flange (Left-type)':['Bottom-Flange (Left-spacing)', df_section['t1'][run_id]+df_section['H'][run_id], df_section['Ref_bot'][run_id], [df_section['Ref_bot'][run_id], df_section['Ref_bot'][run_id]]],
                            'Bottom-Flange (Center-type)':['Bottom-Flange (Center-spacing)',  df_section['t1'][run_id]+df_section['H'][run_id], df_section['Ref_bot'][run_id]+df_section['B4'][run_id], [df_section['Ref_bot'][run_id], df_section['Ref_bot'][run_id]]],
                            'Bottom-Flange (Right-type)':['Bottom-Flange (Right-spacing)',  df_section['t1'][run_id]+df_section['H'][run_id], df_section['Ref_bot'][run_id]+df_section['B4'][run_id]+df_section['B5'][run_id], [df_section['Ref_bot'][run_id], df_section['Ref_bot'][run_id]]],
                            }
        section_area, section_zna, section_yna, section_iyy, section_izz, rib_if_top, rib_if_bot = Includeribs(df_section, df_rib_property, run_id, area_girder, z_girder, y_girder, iyy_girder, izz_girder, dict_position_top, dict_position_bot)

        '''Shear & Torsion調整(僅含箱梁不含懸伸翼版)'''
        ## Shear
        # NOTE 在斜腹鈑時可能不是這樣算，但無可以驗證的例子
        difference_y1 = df_section['Ref_top'][run_id] +df_section['B1'][run_id] -df_section['Ref_bot'][run_id] -df_section['B4'][run_id]
        inclineangle1 = math.atan(difference_y1 /df_section['H'][run_id])
        thicknesswide1 = df_section['tw1'][run_id]/math.cos(inclineangle1)

        difference_y2 = df_section['Ref_top'][run_id] +df_section['B1'][run_id] +df_section['B2'][run_id] -df_section['Ref_bot'][run_id] -df_section['B4'][run_id] -df_section['B5'][run_id]
        inclineangle2 = math.atan(difference_y2 /df_section['H'][run_id])
        thicknesswide2 = df_section['tw2'][run_id]/math.cos(inclineangle2)

        shear_bt = df_section['B2'][run_id] +thicknesswide1 +thicknesswide2
        shear_bb = df_section['B5'][run_id] +thicknesswide1 +thicknesswide2
        shear_t1 = df_section['t1'][run_id]
        shear_t2 = df_section['t2'][run_id]
        shear_h = df_section['H'][run_id]


        section_asy = shear_bt*shear_t1 + shear_bb*shear_t2
        section_asz = shear_h*thicknesswide1 + shear_h*thicknesswide2

        ## Torison
        # 4A^2/(b/t)
        torsion_bt = df_section['B2'][run_id] +thicknesswide1/2 +thicknesswide2/2
        torsion_bb = df_section['B5'][run_id] +thicknesswide1/2 +thicknesswide2/2
        torsion_h = df_section['H'][run_id] +df_section['t1'][run_id]/2 +df_section['t2'][run_id]/2
        torsion_area = (torsion_bt + torsion_bb)*torsion_h/2

        section_ixx = 4*(torsion_area)**2 /(torsion_bt/shear_t1 +torsion_bb/shear_t2 +torsion_h/thicknesswide1 +torsion_h/thicknesswide2)   # 理論上在積分時應該是沿著斜邊H去走路徑，但因為都要處理cos問題，因此直接帶斜邊的tw

        '''Plastic Section Modulus'''
        ## 頂版貢獻
        ### Zyy
        zyy_top = (shear_bt*shear_t1)*abs(section_zna - shear_t1/2)
        ### Zzz
        section_yna_right = df_section['B1'][run_id] +df_section['B2'][run_id] +df_section['B3'][run_id] -section_yna
        zzz_top_left = (section_yna*shear_t1)*abs(section_yna/2)
        zzz_top_right = (section_yna_right*shear_t1)*abs(section_yna_right/2)
        zzz_top = zzz_top_left +zzz_top_right

        ## 底版貢獻
        ### Zyy
        section_zna_bot = df_section['t1'][run_id] +df_section['H'][run_id] +df_section['t2'][run_id] -section_zna
        zyy_bot = (shear_bb*shear_t2)*abs(section_zna_bot - shear_t2/2)
        ### Zzz
        dist_bot_leftna = section_yna -df_section['Ref_bot'][run_id]
        dist_bot_rightna = df_section['B4'][run_id] +df_section['B5'][run_id] +df_section['B6'][run_id] -dist_bot_leftna
        zzz_bot_left = (dist_bot_leftna*shear_t2)*abs(dist_bot_leftna/2)
        zzz_bot_right = (dist_bot_rightna*shear_t2)*abs(dist_bot_rightna/2)
        zzz_bot = zzz_bot_left +zzz_bot_right

        # NOTE 腹鈑貢獻部分，以通常情境計算，未處理NA軸跑進翼版狀況
        ## 腹版1貢獻
        ### Zyy
        dist_zna_w1_1 = section_zna -df_section['t1'][run_id]
        zyy_w1_1 = (dist_zna_w1_1*thicknesswide1)*abs(dist_zna_w1_1/2)
        dist_zna_w1_2 = section_zna_bot -df_section['t2'][run_id]
        zyy_w1_2 = (dist_zna_w1_2*thicknesswide1)*abs(dist_zna_w1_2/2)
        zyy_w1 = zyy_w1_1 +zyy_w1_2
        ### Zzz
        dist_w1 = (df_section['Ref_top'][run_id] +df_section['B1'][run_id] + df_section['Ref_bot'][run_id] +df_section['B4'][run_id])/2 -thicknesswide1/2
        zzz_w1 = (shear_h*thicknesswide1)*abs(section_yna - dist_w1)
        
        ## 腹版2貢獻
        ### Zyy
        dist_zna_w2_1 = section_zna -df_section['t2'][run_id]
        zyy_w2_1 = (dist_zna_w2_1*thicknesswide2)*abs(dist_zna_w2_1/2)
        dist_zna_w2_2 = section_zna_bot -df_section['t2'][run_id]
        zyy_w2_2 = (dist_zna_w2_2*thicknesswide2)*abs(dist_zna_w2_2/2)
        zyy_w2 = zyy_w2_1 +zyy_w2_2
        ### Zzz
        dist_w2 = (df_section['Ref_top'][run_id] +df_section['B1'][run_id] +df_section['B2'][run_id] + df_section['Ref_bot'][run_id] +df_section['B4'][run_id] +df_section['B5'][run_id])/2 +thicknesswide2/2
        zzz_w2 = (shear_h*thicknesswide2)*abs(section_yna - dist_w2)

        ## 頂版加勁版貢獻
        rib_area_top_temp = rib_if_top[0]
        area_rib_top = rib_if_top[1]
        position_rib_top_z = rib_if_top[2]
        position_rib_top_y = rib_if_top[3]
        ### Zyy
        zyy_rib_top = rib_area_top_temp*abs(section_zna - position_rib_top_z)
        ### Zzz
        zzz_rib_top = 0
        for i in range(len(position_rib_top_y)):
            zzz_rib_top = zzz_rib_top +area_rib_top[i]*abs(section_yna -position_rib_top_y[i])

        ## 底版加勁版貢獻
        rib_area_bot_temp = rib_if_bot[0]
        area_rib_bot = rib_if_bot[1]
        position_rib_bot_z = rib_if_bot[2]
        position_rib_bot_y = rib_if_bot[3]
        ### Zyy
        zyy_rib_bot = rib_area_bot_temp*abs(section_zna - position_rib_bot_z)
        ### Zzz
        zzz_rib_bot = 0
        for i in range(len(position_rib_bot_y)):
            zzz_rib_bot = zzz_rib_bot +area_rib_bot[i]*abs(section_yna -position_rib_bot_y[i])

        ## 總和 
        section_zyy = zyy_top +zyy_bot +zyy_w1 +zyy_w2 +zyy_rib_top +zyy_rib_bot
        section_zzz = zzz_top +zzz_bot +zzz_w1 +zzz_w2 +zzz_rib_top +zzz_rib_bot

        '''彙整結果'''
        dict_section['Name'].append(df_section['Name'][run_id])
        dict_section['Area'].append(section_area)
        dict_section['Asy'].append(section_asy)
        dict_section['Asz'].append(section_asz)
        dict_section['Ixx'].append(section_ixx)
        dict_section['Iyy'].append(section_iyy)
        dict_section['Izz'].append(section_izz)
        dict_section['yna_right'].append(section_yna_right)
        dict_section['yna_left'].append(section_yna)
        dict_section['zna_top'].append(section_zna)
        dict_section['zna_bot'].append(section_zna_bot)
        dict_section['Zyy'].append(section_zyy)
        dict_section['Zzz'].append(section_zzz)

    df_section_stlg = pd.DataFrame.from_dict(dict_section)
    df_section_stlg_sap = pd.DataFrame()
    df_section_stlg_sap['Name'] = df_section_stlg['Name']
    df_section_stlg_sap['S2L'] = (df_section_stlg['Izz']/(1E12))/(df_section_stlg['yna_left']/1000)
    df_section_stlg_sap['S2R'] = (df_section_stlg['Izz']/(1E12))/(df_section_stlg['yna_right']/1000)
    df_section_stlg_sap['S3T'] = (df_section_stlg['Iyy']/(1E12))/(df_section_stlg['zna_top']/1000)
    df_section_stlg_sap['S3B'] = (df_section_stlg['Iyy']/(1E12))/(df_section_stlg['zna_bot']/1000)
    df_section_stlg_sap['R22'] = np.sqrt((df_section_stlg['Izz']/(1E12))/(df_section_stlg['Area']/1E6))
    df_section_stlg_sap['R33'] = np.sqrt((df_section_stlg['Iyy']/(1E12))/(df_section_stlg['Area']/1E6))
    df_section_stlg_sap['t3'] = (df_section_stlg['zna_top']/(1000))+(df_section_stlg['zna_bot']/1000)
    df_section_stlg_sap['t2'] = (df_section_stlg['yna_right']/(1000))+(df_section_stlg['yna_left']/1000)
    df_section_stlg_sap['Area'] = df_section_stlg['Area']/1E6
    df_section_stlg_sap['TorsConst'] = df_section_stlg['Ixx']/1E12
    df_section_stlg_sap['As2'] = df_section_stlg['Asz']/1E6
    df_section_stlg_sap['As3'] = df_section_stlg['Asy']/1E6
    df_section_stlg_sap['I22'] = df_section_stlg['Izz']/1E12
    df_section_stlg_sap['I33'] = df_section_stlg['Iyy']/1E12
    df_section_stlg_sap['Z22'] = df_section_stlg['Zzz']/1E9
    df_section_stlg_sap['Z33'] = df_section_stlg['Zyy']/1E9

    # %% 結果輸出
    output_file = os.path.join(outputpath, inputfilename+"_Result.xlsx")
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df_section_stlg_sap.to_excel(writer, sheet_name='Section_SAP', index=False)
        df_section_stlg.to_excel(writer, sheet_name='Section_Midas', index=False)

    print("> 計算結果輸出至 {}".format(inputfilename+"_Result.xlsx"))


def Equivalentwidth(inputfile):
    print('$ 計算等價支間長')
    # %% 讀檔
    (outputpath, filename_temp) = os.path.split(inputfile)
    inputfilename = filename_temp.split(".")[0]
    df_section = pd.read_excel(inputfile, sheet_name='等價支間長', skiprows=[1])

    '''有效寬度分區處理'''
    print('> 處理計算分區')
    df_effective_zone = pd.DataFrame(columns=['Dist', 'Span', 'IE', 'l', 'KxLx', 'KyLy'])
    df_effective_zone['Name'] = df_section['Name']
    df_effective_zone = df_effective_zone.set_index('Name')
    girdercategory = dict(tuple(df_section.groupby('主梁分類')))
    for girder, gdata in girdercategory.items():
        df_effective_info = pd.DataFrame()
        df_effective_info['Name'] = gdata['Name']
        # 找出支承位置與跨度
        support_position_idx = gdata[gdata['支承處'] == 'Y'].index
        support_position = gdata.loc[support_position_idx, '距起點距離']
        support_position = support_position.sort_values().tolist()
        span = np.diff(support_position)

        df_effective_info['Dist'] = gdata['距起點距離']
        # 處裡每個位置span註記分類
        spanlabels = [f'span{i+1}' for i in range(len(support_position)-1)]
        # 先用 cut 分類
        df_effective_info['Span'] = pd.cut(gdata['距起點距離'], bins=support_position, labels=spanlabels, right=False)
        # 判斷是否是 bins 裡的值，並標註柱位
        # df_effective_info['Span'] = np.where(gdata['距起點距離'].isin(support_position), 'C', df_effective_info['Span'])
        df_effective_info['Span'] = gdata['距起點距離'].apply(lambda x: f'C{support_position.index(x) + 1}' if x in support_position else df_effective_info['Span'][gdata['距起點距離'] == x].values[0])

        # IE註記
        controlpoint  = {}
        for i in range(len(support_position)-1):
            controlpoint[i] = []
            start = support_position[i]
            end = support_position[i + 1]
            span_c = end - start

            # 第一個 span
            if i == 0:
                controlpoint[i].append(start)
                controlpoint[i].append(start + 0.8 * span_c)
                controlpoint[i].append(end)
            # 最後一個 span
            elif i == len(support_position) - 2:
                controlpoint[i].append(start)
                controlpoint[i].append(start + 0.2 * span_c)
                controlpoint[i].append(end)
            # 中間的 span
            else:
                controlpoint[i].append(start)
                controlpoint[i].append(start + 0.2 * span_c)
                controlpoint[i].append(start + 0.8 * span_c)
                controlpoint[i].append(end)

        df_effective_info = df_effective_info.set_index('Name')
        df_effective_info['IE'] = None
        df_effective_info['l'] = None
        df_effective_info['KxLx'] = None
        df_effective_info['KyLy'] = None
        for index, data in df_effective_info.iterrows():
            # 處理第一跨點位
            if data['Span'] == 'span1':
                if data['Dist'] <= controlpoint[0][1]:
                    df_effective_info.loc[index, 'IE'] = 1
                    df_effective_info.loc[index, 'l'] = 0.8*span[0]
                    df_effective_info.loc[index, 'KxLx'] = span[0]
                    df_effective_info.loc[index, 'KyLy'] = span[0]
                else:
                    df_effective_info.loc[index, 'IE'] = 2
                    df_effective_info.at[index, 'l'] = [data['Dist'], 
                                                        controlpoint[0][1], 
                                                        controlpoint[0][2],
                                                        0.8*span[0],
                                                        0.2*(span[0]+span[1])
                                                        ]
                    df_effective_info.loc[index, 'KxLx'] = span[0]
                    df_effective_info.loc[index, 'KyLy'] = span[0]
            
            # 處理最後一跨點位
            elif data['Span'] == 'span'+str(len(span)):
                if data['Dist'] >= controlpoint[len(span)-1][1]:
                    df_effective_info.loc[index, 'IE'] = 1
                    df_effective_info.loc[index, 'l'] = 0.8*span[-1]
                    df_effective_info.loc[index, 'KxLx'] = span[-1]
                    df_effective_info.loc[index, 'KyLy'] = span[-1]
                else:
                    df_effective_info.loc[index, 'IE'] = 4
                    df_effective_info.at[index, 'l'] = [data['Dist'], 
                                                        controlpoint[len(span)-1][0], 
                                                        controlpoint[len(span)-1][1],
                                                        0.2*(span[-2]+span[-1]),
                                                        0.8*span[-1]
                                                        ]
                    df_effective_info.loc[index, 'KxLx'] = span[-1]
                    df_effective_info.loc[index, 'KyLy'] = span[-1]
            
            # 處理第一墩
            elif data['Span'] == 'C1':
                df_effective_info.loc[index, 'IE'] = 1
                df_effective_info.loc[index, 'l'] = 0.8*span[0]
                df_effective_info.loc[index, 'KxLx'] = span[0]
                df_effective_info.loc[index, 'KyLy'] = span[0]

            # 處理最後一墩
            elif data['Span'] == 'C'+str(len(span)+1):
                df_effective_info.loc[index, 'IE'] = 1
                df_effective_info.loc[index, 'l'] = 0.8*span[-1]
                df_effective_info.loc[index, 'KxLx'] = span[-1]
                df_effective_info.loc[index, 'KyLy'] = span[-1]

            # 處理其他墩
            elif 'C' in data['Span']:
                df_effective_info.loc[index, 'IE'] = 3
                span_pre = int(data['Span'].split('C')[-1])-2
                span_post = int(data['Span'].split('C')[-1])-1
                df_effective_info.loc[index, 'l'] = 0.2*(span[span_pre]+span[span_post])
                df_effective_info.loc[index, 'KxLx'] = 0.5*(span[span_pre]+span[span_post])
                df_effective_info.loc[index, 'KyLy'] = 0.5*(span[span_pre]+span[span_post])

            # 處理剩餘跨
            else:
                span_id = int(data['Span'].split('span')[-1])-1
                if data['Dist'] <= controlpoint[span_id][1]:
                    df_effective_info.loc[index, 'IE'] = 4
                    df_effective_info.at[index, 'l'] = [data['Dist'], 
                                                        controlpoint[span_id][0], 
                                                        controlpoint[span_id][1],
                                                        0.2*(span[span_id-1]+span[span_id]),
                                                        0.6*span[span_id]
                                                        ]
                    df_effective_info.loc[index, 'KxLx'] = span[span_id]
                    df_effective_info.loc[index, 'KyLy'] = span[span_id]
                elif data['Dist'] >= controlpoint[span_id][2]:
                    df_effective_info.loc[index, 'IE'] = 2
                    df_effective_info.at[index, 'l'] = [data['Dist'], 
                                                        controlpoint[span_id][2], 
                                                        controlpoint[span_id][3],
                                                        0.6*span[span_id],
                                                        0.2*(span[span_id]+span[span_id+1]),
                                                        ]  
                    df_effective_info.loc[index, 'KxLx'] = span[span_id]
                    df_effective_info.loc[index, 'KyLy'] = span[span_id]         
                else:
                    df_effective_info.loc[index, 'IE'] = 5
                    df_effective_info.loc[index, 'l'] = 0.6*span[span_id]
                    df_effective_info.loc[index, 'KxLx'] = span[span_id]
                    df_effective_info.loc[index, 'KyLy'] = span[span_id]

        df_effective_zone.update(df_effective_info)

    # %% 結果輸出
    output_file = os.path.join(outputpath, inputfilename+"_EqualSpan.xlsx")
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df_effective_zone.to_excel(writer, sheet_name='等價支間長')

    print("> 計算結果輸出至 {}".format(inputfilename+"_EqualSpan.xlsx"))


def Effectivesection(inputfile):
    """計算鋼床鈑有效斷面

    Args:
        inputfile (str): 輸入參數的excel
    """
    print("$ 執行有效斷面計算。")
    # %% 讀檔
    (outputpath, filename_temp) = os.path.split(inputfile)
    inputfilename = filename_temp.split(".")[0]

    df_section = pd.read_excel(inputfile, sheet_name='鋼床鈑', skiprows=[1])
    df_ribs = pd.read_excel(inputfile, sheet_name='加勁鈑')

    # %% 計算加勁鈑
    dict_rib_property = Ribproperty(df_ribs)
    df_rib_property = pd.DataFrame.from_dict(dict_rib_property)
    df_rib_property = df_rib_property.set_index('Name')  

    # %% 有效斷面計算
    dict_effectivesection = {'Name':[],
                            'EffectiveWide(top)':[],
                            'Formula(top)':[],
                            'b/l(top)':[],
                            'EffectiveWide(bot)':[],
                            'Formula(bot)':[],
                            'b/l(bot)':[],
                            'EffectiveWide(web)':[],
                            'Formula(web)':[],
                            'b/l(web)':[],
                            'Area_y':[],
                            'Iyy':[],
                            'Area_z':[],
                            'Izz':[],
                            'yna_right':[],
                            'yna_left':[],
                            'zna_top':[],
                            'zna_bot':[],
                            }
    for run_id in range(len(df_section)):
        '''主要箱梁有效斷面性質'''
        difference_y1 = df_section['Ref_top'][run_id] +df_section['B1'][run_id] -df_section['Ref_bot'][run_id] -df_section['B4'][run_id]
        inclineangle1 = math.atan(difference_y1 /df_section['H'][run_id])
        thicknesswide1 = df_section['tw1'][run_id]/math.cos(inclineangle1)

        difference_y2 = df_section['Ref_top'][run_id] +df_section['B1'][run_id] +df_section['B2'][run_id] -df_section['Ref_bot'][run_id] -df_section['B4'][run_id] -df_section['B5'][run_id]
        inclineangle2 = math.atan(difference_y2 /df_section['H'][run_id])
        thicknesswide2 = df_section['tw2'][run_id]/math.cos(inclineangle2)

        # 有效幅計算
        # NOTE: 代入公式的b採淨間距
        b_top_1 = df_section['B1'][run_id] -thicknesswide1
        b_top_2 = df_section['B2'][run_id]/2
        b_top_3 = df_section['B3'][run_id] -thicknesswide2
        b_bot_1 = df_section['B4'][run_id] -thicknesswide1
        b_bot_2 = df_section['B5'][run_id]/2
        b_bot_3 = df_section['B6'][run_id] -thicknesswide2
        b_h = df_section['H'][run_id]/2

        ie = df_section['IE'][run_id]
        lei = df_section['l'][run_id]
        leo = df_section['l'][run_id]

        def effectiveflange(IE, B, LE):
            match IE:
                case 1|5:
                    # (11.3.1)
                    note_formula = '11.3.1'
                    note_bl = B/LE
                    if B/LE <= 0.05:
                        lambda_e = B
                    elif B/LE <= 0.3:
                        lambda_e = (1.1 - 2*(B/LE))*B
                    else:
                        lambda_e = 0.15*LE

                case 3:
                    # (11.3.2)
                    note_formula = '11.3.2'
                    note_bl = B/LE
                    if B/LE <= 0.02:
                        lambda_e = B
                    elif B/LE <= 0.3:
                        lambda_e = (1.06 - 3.2*(B/LE) + 4.5*(B/LE)**2)*B
                    else:
                        lambda_e = 0.15*LE
                    
                case 2:
                    # 使用兩端之有效幅作線性內插
                    note_formula = '內插區'
                    note_bl = '內插'
                    insert_info = ast.literal_eval(LE)
                    dist_current = insert_info[0]
                    dist_start = insert_info[1]
                    dist_end = insert_info[2]
                    l_start = insert_info[3]
                    l_end = insert_info[4]

                    if B/l_start <= 0.05:
                        lambda_start = B
                    elif B/l_start <= 0.3:
                        lambda_start = (1.1 - 2*(B/l_start))*B
                    else:
                        lambda_start = 0.15*l_start
                    
                    if B/l_end <= 0.02:
                        lambda_end = B
                    elif B/l_end <= 0.3:
                        lambda_end = (1.06 - 3.2*(B/l_end) + 4.5*(B/l_end)**2)*B
                    else:
                        lambda_end = 0.15*l_end

                    lambda_e = lambda_start + (dist_current - dist_start)*(lambda_end - lambda_start)/(dist_end - dist_start)

                case 4:
                    # 使用兩端之有效幅作線性內插
                    note_formula = '內插區'
                    note_bl = '內插'
                    insert_info = ast.literal_eval(LE)
                    dist_current = insert_info[0]
                    dist_start = insert_info[1]
                    dist_end = insert_info[2]
                    l_start = insert_info[3]
                    l_end = insert_info[4]

                    if B/l_start <= 0.02:
                        lambda_start = B
                    elif B/l_end <= 0.3:
                        lambda_start = (1.06 - 3.2*(B/l_start) + 4.5*(B/l_start)**2)*B
                    else:
                        lambda_start = 0.15*l_start

                    if B/l_end <= 0.05:
                        lambda_end = B
                    elif B/l_end <= 0.3:
                        lambda_end = (1.1 - 2*(B/l_end))*B
                    else:
                        lambda_end = 0.15*l_end

                    lambda_e = lambda_start + (dist_current - dist_start)*(lambda_end - lambda_start)/(dist_end - dist_start)

            return lambda_e, note_formula, note_bl

        lambda_top1, note_formula_top1, note_bl_top1 = effectiveflange(ie, b_top_1, lei)
        lambda_top2, note_formula_top2, note_bl_top2 = effectiveflange(ie, b_top_2, lei)
        lambda_top3, note_formula_top3, note_bl_top3 = effectiveflange(ie, b_top_3, lei)
        lambda_bot1, note_formula_bot1, note_bl_bot1 = effectiveflange(ie, b_bot_1, lei)
        lambda_bot2, note_formula_bot2, note_bl_bot2 = effectiveflange(ie, b_bot_2, lei)
        lambda_bot3, note_formula_bot3, note_bl_bot3 = effectiveflange(ie, b_bot_3, lei)
        lambda_h, note_formula_h, note_bl_h = effectiveflange(ie, b_h, leo)

        ## 有效斷面控制點
        top_e1 = df_section['Ref_top'][run_id] + df_section['B1'][run_id] - thicknesswide1 - lambda_top1
        top_e2 = df_section['Ref_top'][run_id] + df_section['B1'][run_id] + lambda_top2
        top_e3 = df_section['Ref_top'][run_id] + df_section['B1'][run_id] + df_section['B2'][run_id] - lambda_top2
        top_e4 = df_section['Ref_top'][run_id] + df_section['B1'][run_id] + df_section['B2'][run_id] + thicknesswide2 + lambda_top3
        bot_e1 = df_section['Ref_bot'][run_id] + df_section['B4'][run_id] - thicknesswide1 - lambda_bot1
        bot_e2 = df_section['Ref_bot'][run_id] + df_section['B4'][run_id] + lambda_bot2
        bot_e3 = df_section['Ref_bot'][run_id] + df_section['B4'][run_id] + df_section['B5'][run_id] - lambda_bot2
        bot_e4 = df_section['Ref_bot'][run_id] + df_section['B4'][run_id] + df_section['B5'][run_id] + thicknesswide2 + lambda_bot3
        h_e1 = df_section['t1'][run_id] + lambda_h
        h_e2 = df_section['t1'][run_id] + df_section['H'][run_id] - lambda_h

        # 主梁有效斷面計算
        # NOTE: 座標系統名稱與MIDAS一致(以前反卡氏xy)，以便對照全斷面結果
        # 對y軸
        area_ye, z_e, y_, iyy_e, izz_ = Girdersection(lambda_top1 + thicknesswide1 + lambda_top2 + lambda_top2 + thicknesswide2 + lambda_top3, 
                                                        df_section['t1'][run_id], 
                                                        lambda_bot1 + thicknesswide1 + lambda_bot2 + lambda_bot2 + thicknesswide2 + lambda_bot3,
                                                        df_section['t2'][run_id], 
                                                        df_section['H'][run_id], 
                                                        thicknesswide1, 
                                                        thicknesswide2, 
                                                        df_section['Ref_top'][run_id], 
                                                        df_section['Ref_bot'][run_id], 
                                                        df_section['B1'][run_id], 
                                                        df_section['B2'][run_id], 
                                                        df_section['B4'][run_id], 
                                                        df_section['B5'][run_id]
                                                        )

        # 對z軸
        area_ze, z_, y_e, iyy_, izz_e = Girdersection(df_section['B1'][run_id] +df_section['B2'][run_id] +df_section['B3'][run_id], 
                                                        df_section['t1'][run_id], 
                                                        df_section['B4'][run_id] +df_section['B5'][run_id] +df_section['B6'][run_id],
                                                        df_section['t2'][run_id], 
                                                        lambda_h*2, 
                                                        thicknesswide1, 
                                                        thicknesswide2, 
                                                        df_section['Ref_top'][run_id], 
                                                        df_section['Ref_bot'][run_id], 
                                                        df_section['B1'][run_id], 
                                                        df_section['B2'][run_id], 
                                                        df_section['B4'][run_id], 
                                                        df_section['B5'][run_id]
                                                        )

        '''加入加勁鈑'''
        # 對y軸
        dict_ye_top = {'Top-Flange (Left-type)':['Top-Flange (Left-spacing)', df_section['t1'][run_id], df_section['Ref_top'][run_id], [df_section['Ref_top'][run_id], top_e1]],
                        'Top-Flange (Center-type)':['Top-Flange (Center-spacing)', df_section['t1'][run_id], df_section['Ref_top'][run_id]+df_section['B1'][run_id], [top_e2, top_e3]],
                        'Top-Flange (Right-type)':['Top-Flange (Right-spacing)', df_section['t1'][run_id], df_section['Ref_top'][run_id]+df_section['B1'][run_id]+df_section['B2'][run_id], [top_e4, df_section['Ref_top'][run_id]+df_section['B1'][run_id]+df_section['B2'][run_id]+df_section['B3'][run_id]]],
                        }
        dict_ye_bot = {'Bottom-Flange (Left-type)':['Bottom-Flange (Left-spacing)', df_section['t1'][run_id]+df_section['H'][run_id], df_section['Ref_bot'][run_id], [df_section['Ref_bot'][run_id], bot_e1]],
                        'Bottom-Flange (Center-type)':['Bottom-Flange (Center-spacing)',  df_section['t1'][run_id]+df_section['H'][run_id], df_section['Ref_bot'][run_id]+df_section['B4'][run_id], [bot_e2, bot_e3]],
                        'Bottom-Flange (Right-type)':['Bottom-Flange (Right-spacing)',  df_section['t1'][run_id]+df_section['H'][run_id], df_section['Ref_bot'][run_id]+df_section['B4'][run_id]+df_section['B5'][run_id], [bot_e4, df_section['Ref_bot'][run_id]+df_section['B4'][run_id]+df_section['B5'][run_id]+df_section['B6'][run_id]]],
                        }
        section_area_ye, section_zna_ye, section_yna_ye, section_iyy_ye, section_izz_ye, rib_info_top_ye, rib_info_bot_ye = Includeribs(df_section, df_rib_property, run_id, area_ye, z_e, y_, iyy_e, izz_, dict_ye_top, dict_ye_bot)

        # 對z軸
        dict_ze_top = {'Top-Flange (Left-type)':['Top-Flange (Left-spacing)', df_section['t1'][run_id], df_section['Ref_top'][run_id], [df_section['Ref_top'][run_id], df_section['Ref_top'][run_id]]],
                        'Top-Flange (Center-type)':['Top-Flange (Center-spacing)', df_section['t1'][run_id], df_section['Ref_top'][run_id]+df_section['B1'][run_id], [top_e2, top_e2]],
                        'Top-Flange (Right-type)':['Top-Flange (Right-spacing)', df_section['t1'][run_id], df_section['Ref_top'][run_id]+df_section['B1'][run_id]+df_section['B2'][run_id], [top_e4, top_e4]],
                        }
        dict_ze_bot = {'Bottom-Flange (Left-type)':['Bottom-Flange (Left-spacing)', df_section['t1'][run_id]+df_section['H'][run_id], df_section['Ref_bot'][run_id], [df_section['Ref_bot'][run_id], df_section['Ref_bot'][run_id]]],
                        'Bottom-Flange (Center-type)':['Bottom-Flange (Center-spacing)',  df_section['t1'][run_id]+df_section['H'][run_id], df_section['Ref_bot'][run_id]+df_section['B4'][run_id], [bot_e2, bot_e2]],
                        'Bottom-Flange (Right-type)':['Bottom-Flange (Right-spacing)',  df_section['t1'][run_id]+df_section['H'][run_id], df_section['Ref_bot'][run_id]+df_section['B4'][run_id]+df_section['B5'][run_id], [bot_e4, bot_e4]],
                        }
        section_area_ze, section_zna_ze, section_yna_ze, section_iyy_ze, section_izz_ze, rib_info_top_ze, rib_info_bot_ze = Includeribs(df_section, df_rib_property, run_id, area_ze, z_, y_e, iyy_, izz_e, dict_ze_top, dict_ze_bot)

        section_zna_bot = df_section['t1'][run_id] +df_section['H'][run_id] +df_section['t2'][run_id] -section_zna_ye
        section_yna_right = df_section['B1'][run_id] +df_section['B2'][run_id] +df_section['B3'][run_id] -section_yna_ze

        '''彙整結果'''
        dict_effectivesection['Name'].append(df_section['Name'][run_id])
        dict_effectivesection['EffectiveWide(top)'].append([lambda_top1, lambda_top2, lambda_top3])
        dict_effectivesection['Formula(top)'].append([note_formula_top1, note_formula_top2, note_formula_top3])
        dict_effectivesection['b/l(top)'].append([note_bl_top1, note_bl_top2, note_bl_top3])
        dict_effectivesection['EffectiveWide(bot)'].append([lambda_bot1, lambda_bot2, lambda_bot3])
        dict_effectivesection['Formula(bot)'].append([note_formula_bot1, note_formula_bot2, note_formula_bot3])
        dict_effectivesection['b/l(bot)'].append([note_bl_bot1, note_bl_bot2, note_bl_bot3])
        dict_effectivesection['EffectiveWide(web)'].append([lambda_h])
        dict_effectivesection['Formula(web)'].append([note_formula_h])
        dict_effectivesection['b/l(web)'].append([note_bl_h])
        dict_effectivesection['Area_y'].append(section_area_ye)
        dict_effectivesection['Iyy'].append(section_iyy_ye)
        dict_effectivesection['zna_top'].append(section_zna_ye)
        dict_effectivesection['zna_bot'].append(section_zna_bot)
        dict_effectivesection['Area_z'].append(section_area_ze)
        dict_effectivesection['Izz'].append(section_izz_ze)
        dict_effectivesection['yna_right'].append(section_yna_right)
        dict_effectivesection['yna_left'].append(section_yna_ze)

    df_effectivesection_stlg = pd.DataFrame.from_dict(dict_effectivesection)
    df_effectivesection_stlg_sap = pd.DataFrame()
    df_effectivesection_stlg_sap['Name'] = df_effectivesection_stlg['Name']
    df_effectivesection_stlg_sap['Axe'] = df_effectivesection_stlg['Area_y']/1E6
    df_effectivesection_stlg_sap['Ixe'] = df_effectivesection_stlg['Iyy']/1E12
    df_effectivesection_stlg_sap['SxeT'] = (df_effectivesection_stlg['Iyy']/(1E12))/(df_effectivesection_stlg['zna_top']/1000)
    df_effectivesection_stlg_sap['SxeB'] = (df_effectivesection_stlg['Iyy']/(1E12))/(df_effectivesection_stlg['zna_bot']/1000)
    df_effectivesection_stlg_sap['Aye'] = df_effectivesection_stlg['Area_z']/1E6
    df_effectivesection_stlg_sap['Iye'] = df_effectivesection_stlg['Izz']/1E12
    df_effectivesection_stlg_sap['SyeL'] = (df_effectivesection_stlg['Izz']/(1E12))/(df_effectivesection_stlg['yna_left']/1000)
    df_effectivesection_stlg_sap['SyeR'] = (df_effectivesection_stlg['Izz']/(1E12))/(df_effectivesection_stlg['yna_right']/1000)

    # %% 結果輸出
    output_file = os.path.join(outputpath, inputfilename+"_EffectiveSec.xlsx")
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df_effectivesection_stlg_sap.to_excel(writer, sheet_name='Section_SDB', index=False)
        df_effectivesection_stlg.to_excel(writer, sheet_name='Section_Calculation', index=False)

    print("> 計算結果輸出至 {}".format(inputfilename+"_EffectiveSec.xlsx"))


def Allowablestress(inputfile):
    """計算鋼床鈑容許應力

    Args:
        inputfile (str): 輸入參數的excel
    """
    print("$ 執行容許應力計算。")
    # %% 讀檔
    (outputpath, filename_temp) = os.path.split(inputfile)
    inputfilename = filename_temp.split(".")[0]
    df_section = pd.read_excel(inputfile, sheet_name='鋼床鈑', skiprows=[1])

    section_file = os.path.join(outputpath, inputfilename+"_Result.xlsx")
    try:
        df_sectionproperty = pd.read_excel(section_file, sheet_name='Section_SAP')
    except:
        print("> [Warning]: 請先執行斷面計算生成 [輸入檔名稱]_Result.xlsx")
        sys.exit()

    dict_allowablestress = {'Name':[],
                            'FbxB':[],
                            'FbxT':[],
                            'Fv':[],
                            'Fa':[],
                            'Fby':[],
                            'Fex':[],
                            'Fey':[],
                            'KyLy/ry':[],
                            'KxLx/rx':[],
                            }
    for run_id in range(len(df_section)):
        '''Fa'''
        Es = df_section['E'][run_id]
        Fy = df_section['Fy'][run_id]
        Kx = df_section['KxLx'][run_id]/(df_sectionproperty['R33'][run_id]*1000)
        Ky = df_section['KyLy'][run_id]/(df_sectionproperty['R22'][run_id]*1000)
        FS = 2.12

        Cc = math.sqrt(2*math.pow(math.pi,2)*Es/Fy)
        k = Kx if Kx >= Ky else Ky

        if k <= Cc:
            Fa = Fy/FS*(1-(math.pow(k,2)*Fy/(4*math.pow(math.pi,2)*Es)))
        else:
            Fa = math.pow(math.pi,2)*Es/(FS*math.pow(k,2))

        '''Fb'''
        def Fb(tf, B, SP, Fy, Nrib, w):
            k = min((math.pow(1+math.pow(SP/B,2),2) + 87.3) / (math.pow(Nrib+1,2)*math.pow(SP/B,2)*(1+0.1*(Nrib+1))) , 4)
            wt = w/tf

            if wt <= 813.96*math.sqrt(k/Fy):
                Fb = 0.55*Fy
            elif wt <= min(1763*math.sqrt(k/Fy), 60):
                Fb = 0.55*Fy - 0.224*Fy*(1-math.sin(0.5*math.pi*(1763*math.sqrt(k)-wt*math.sqrt(Fy))/(949*math.sqrt(k))))
            elif wt > 1763*math.sqrt(k/Fy) and wt <= 60:
                Fb = 1.01408*k/math.pow(wt,2)*1E6
            else:
                print('*[Warning]: w/t > 60')
                Fb = 'N/A'
            
            return Fb
        
        # NOTE:應該間距越大強度越低，這裡採輸入間距中最大值
        rib_spacing = re.split(r'\s*,\s*', df_section['Top-Flange (Center-spacing)'][run_id])
        FbxT = Fb(df_section['t1'][run_id], df_section['B2'][run_id], df_section['SP'][run_id], df_section['Fy'][run_id], len(rib_spacing), float(max(rib_spacing)))

        rib_spacing = re.split(r'\s*,\s*', df_section['Bottom-Flange (Center-spacing)'][run_id])
        FbxB = Fb(df_section['t2'][run_id], df_section['B5'][run_id], df_section['SP'][run_id], df_section['Fy'][run_id], len(rib_spacing), float(max(rib_spacing)))

        '''Fv'''
        tw = min(df_section['tw1'][run_id], df_section['tw2'][run_id])
        H = df_section['H'][run_id]
        D0 = df_section['D0'][run_id]
        Fy = df_section['Fy'][run_id]

        # SDB求C方法
        C = min(1.55*1E7*(1+math.pow(H/D0,2))/(Fy*math.pow(H/tw,2)), 1)
        # 規範標準解法
        #
        # k = 5 + 5/math.pow(D0/H,2)
        # if H/tw < 1590*math.sqrt(k/Fy):
        #     C = 1
        # elif H/tw <= 1990*math.sqrt(k/Fy):
        #     C = 1590*math.sqrt(k/Fy)/(H/tw)
        # else:
        #     C = 3.17*1E6*k/(math.pow(H/tw,2)*Fy)
        #
        Fv = (Fy/3) * (C + 0.87*(1-C)/math.sqrt(1+math.pow(D0/H,2)))

        '''Fby'''
        Fby = 0.55*df_section['Fy'][run_id]

        '''Fe'''
        Es = df_section['E'][run_id]
        Fy = df_section['Fy'][run_id]
        Kx = df_section['KxLx'][run_id]/(df_sectionproperty['R33'][run_id]*1000)
        Ky = df_section['KyLy'][run_id]/(df_sectionproperty['R22'][run_id]*1000)
        FS = 2.12

        Fex = math.pow(math.pi,2)*Es/(FS*math.pow(Kx,2))
        Fey = math.pow(math.pi,2)*Es/(FS*math.pow(Ky,2))


        '''彙整結果'''
        dict_allowablestress['Name'].append(df_section['Name'][run_id])
        dict_allowablestress['FbxB'].append(FbxB)
        dict_allowablestress['FbxT'].append(FbxT)
        dict_allowablestress['Fv'].append(Fv)
        dict_allowablestress['Fa'].append(Fa)
        dict_allowablestress['Fby'].append(Fby)
        dict_allowablestress['Fex'].append(Fex)
        dict_allowablestress['Fey'].append(Fey)
        dict_allowablestress['KyLy/ry'].append(Ky)
        dict_allowablestress['KxLx/rx'].append(Kx)


    df_allowablestress = pd.DataFrame.from_dict(dict_allowablestress)

    # %% 結果輸出
    output_file = os.path.join(outputpath, inputfilename+"_AllowStress.xlsx")
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df_allowablestress.to_excel(writer, sheet_name='AllowableStress_SAP', index=False)

    print("> 計算結果輸出至 {}".format(inputfilename+"_AllowStress.xlsx"))



def Plotsectiondxf(inputfile):
    print("$ 執行斷面DXF繪製。")
    # %% 讀檔
    (outputpath, filename_temp) = os.path.split(inputfile)
    inputfilename = filename_temp.split(".")[0]

    df_section = pd.read_excel(inputfile, sheet_name='鋼床鈑', skiprows=[1])
    df_ribs = pd.read_excel(inputfile, sheet_name='加勁鈑')
    df_rib_property = pd.DataFrame.from_dict(df_ribs)
    df_rib_property = df_rib_property.set_index('Name') 

    drawing_spacing = 0
    # create a new DXF R2010 document
    doc = ezdxf.new()
    # add new entities to the modelspace
    msp = doc.modelspace()
    if 'Section' not in doc.layers:
        doc.layers.new(name='Section')
    if 'Annotation' not in doc.layers:
        doc.layers.new(name='Annotation')
    text_style = "FontStyle"
    if text_style not in doc.styles:
        doc.styles.new(name=text_style, dxfattribs={"font" : "OpenSans-Regular.ttf"}) 
    for run_id in range(len(df_section)):
        '''主要箱梁斷面'''
        # 處理腹板厚
        H = df_section['H'][run_id]
        tw1 = df_section['tw1'][run_id]
        tw2 = df_section['tw2'][run_id]
        difference_y1 = df_section['Ref_top'][run_id] +df_section['B1'][run_id] -df_section['Ref_bot'][run_id] -df_section['B4'][run_id]
        inclineangle1 = math.atan(difference_y1 /H)
        thicknesswide1 = tw1/math.cos(inclineangle1)
        difference_y2 = df_section['Ref_top'][run_id] +df_section['B1'][run_id] +df_section['B2'][run_id] -df_section['Ref_bot'][run_id] -df_section['B4'][run_id] -df_section['B5'][run_id]
        inclineangle2 = math.atan(difference_y2 /H)
        thicknesswide2 = tw2/math.cos(inclineangle2)

        # built the control points of the top flange
        pt1 = (drawing_spacing, 0)
        pt2 = (drawing_spacing+df_section['B1'][run_id]+df_section['B2'][run_id]+df_section['B3'][run_id], 0)
        pt3 = (drawing_spacing, -df_section['t1'][run_id])
        pt4 = (drawing_spacing+df_section['B1'][run_id]+df_section['B2'][run_id]+df_section['B3'][run_id], -df_section['t1'][run_id])
        # built the control points of the botttom flange
        bot_cp = df_section['Ref_bot'][run_id] - df_section['Ref_top'][run_id]
        pb1 = (drawing_spacing+bot_cp, -df_section['t1'][run_id]-H)
        pb2 = (drawing_spacing+bot_cp+df_section['B4'][run_id]+df_section['B5'][run_id]+df_section['B6'][run_id], -df_section['t1'][run_id]-H)
        pb3 = (drawing_spacing+bot_cp, -df_section['t1'][run_id]-H-df_section['t2'][run_id])
        pb4 = (drawing_spacing+bot_cp+df_section['B4'][run_id]+df_section['B5'][run_id]+df_section['B6'][run_id], -df_section['t1'][run_id]-H-df_section['t2'][run_id])
        # built the control points of the left web
        pl1 = (drawing_spacing+df_section['B1'][run_id]-thicknesswide1,  -df_section['t1'][run_id])
        pl2 = (drawing_spacing+df_section['B1'][run_id],  -df_section['t1'][run_id])
        pl3 = (drawing_spacing+bot_cp+df_section['B4'][run_id]-thicknesswide1,  -df_section['t1'][run_id]-H)
        pl4 = (drawing_spacing+bot_cp+df_section['B4'][run_id],  -df_section['t1'][run_id]-H)
        # built the control points of the right web
        pr1 = (drawing_spacing+df_section['B1'][run_id]+df_section['B2'][run_id]-thicknesswide2, -df_section['t1'][run_id])
        pr2 = (drawing_spacing+df_section['B1'][run_id]+df_section['B2'][run_id],  -df_section['t1'][run_id])
        pr3 = (drawing_spacing+bot_cp+df_section['B4'][run_id]+df_section['B5'][run_id]-thicknesswide2,  -df_section['t1'][run_id]-H)
        pr4 = (drawing_spacing+bot_cp+df_section['B4'][run_id]+df_section['B5'][run_id],  -df_section['t1'][run_id]-H)

        # draw section
        msp.add_line(pt1, pt2, dxfattribs={"layer": "Section"})
        msp.add_line(pt1, pt3, dxfattribs={"layer": "Section"})
        msp.add_line(pt2, pt4, dxfattribs={"layer": "Section"})
        msp.add_line(pt3, pt4, dxfattribs={"layer": "Section"})
        msp.add_line(pb1, pb2, dxfattribs={"layer": "Section"})
        msp.add_line(pb1, pb3, dxfattribs={"layer": "Section"})
        msp.add_line(pb2, pb4, dxfattribs={"layer": "Section"})
        msp.add_line(pb3, pb4, dxfattribs={"layer": "Section"})
        msp.add_line(pl1, pl2, dxfattribs={"layer": "Section"})
        msp.add_line(pl1, pl3, dxfattribs={"layer": "Section"})
        msp.add_line(pl2, pl4, dxfattribs={"layer": "Section"})
        msp.add_line(pl3, pl4, dxfattribs={"layer": "Section"})
        msp.add_line(pr1, pr2, dxfattribs={"layer": "Section"})
        msp.add_line(pr1, pr3, dxfattribs={"layer": "Section"})
        msp.add_line(pr2, pr4, dxfattribs={"layer": "Section"})
        msp.add_line(pr3, pr4, dxfattribs={"layer": "Section"})

        '''加入加勁鈑'''
        dict_position_top = {'Top-Flange (Left-type)':['Top-Flange (Left-spacing)', df_section['t1'][run_id], df_section['Ref_top'][run_id], [df_section['Ref_top'][run_id], df_section['Ref_top'][run_id]]],
                            'Top-Flange (Center-type)':['Top-Flange (Center-spacing)', df_section['t1'][run_id], df_section['Ref_top'][run_id]+df_section['B1'][run_id], [df_section['Ref_top'][run_id], df_section['Ref_top'][run_id]]],
                            'Top-Flange (Right-type)':['Top-Flange (Right-spacing)', df_section['t1'][run_id], df_section['Ref_top'][run_id]+df_section['B1'][run_id]+df_section['B2'][run_id], [df_section['Ref_top'][run_id], df_section['Ref_top'][run_id]]],
                            }
        dict_position_bot = {'Bottom-Flange (Left-type)':['Bottom-Flange (Left-spacing)', df_section['t1'][run_id]+df_section['H'][run_id], df_section['Ref_bot'][run_id], [df_section['Ref_bot'][run_id], df_section['Ref_bot'][run_id]]],
                            'Bottom-Flange (Center-type)':['Bottom-Flange (Center-spacing)',  df_section['t1'][run_id]+df_section['H'][run_id], df_section['Ref_bot'][run_id]+df_section['B4'][run_id], [df_section['Ref_bot'][run_id], df_section['Ref_bot'][run_id]]],
                            'Bottom-Flange (Right-type)':['Bottom-Flange (Right-spacing)',  df_section['t1'][run_id]+df_section['H'][run_id], df_section['Ref_bot'][run_id]+df_section['B4'][run_id]+df_section['B5'][run_id], [df_section['Ref_bot'][run_id], df_section['Ref_bot'][run_id]]],
                            }
        # deal with top flange
        for key, item in dict_position_top.items():
            if not pd.isna(df_section[key][run_id]):
                ### 單獨rib斷面性質提取
                rib_type = df_section[key][run_id]
                ### 提取ribs間距
                rib_spacing = re.split(r'\s*,\s*', df_section[item[0]][run_id])
                rib_num = len(rib_spacing)

                ### 參考位置定位與初始化
                rib_z_level = item[1]
                rib_y_level = item[2]
                rib_dist_y = 0
                rib_area_top_temp = 0
                ### 加勁版繪製
                for rr in range(len(rib_spacing)):
                    #### rib座標位置定出
                    set_z = -rib_z_level
                    rib_dist_y = rib_dist_y + float(rib_spacing[rr])
                    set_y = rib_y_level + rib_dist_y

                    if df_rib_property['Type'][rib_type] == 'Flat':
                        rH = df_rib_property['H/H/H'][rib_type]
                        rB = df_rib_property['B/B/B1'][rib_type]

                        rp1 = (drawing_spacing+set_y-rB/2, set_z)
                        rp2 = (drawing_spacing+set_y+rB/2, set_z)
                        rp3 = (drawing_spacing+set_y-rB/2, set_z-rH)
                        rp4 = (drawing_spacing+set_y+rB/2, set_z-rH)
                        msp.add_line(rp1, rp2, dxfattribs={"layer": "Section"})
                        msp.add_line(rp1, rp3, dxfattribs={"layer": "Section"})
                        msp.add_line(rp2, rp4, dxfattribs={"layer": "Section"})
                        msp.add_line(rp3, rp4, dxfattribs={"layer": "Section"})

                    elif df_rib_property['Type'][rib_type] == 'Tee':
                        bt = df_rib_property['B/B/B1'][rib_type] 
                        tft = df_rib_property['/tf/t'][rib_type] 
                        ht = df_rib_property['H/H/H'][rib_type] 
                        twt = df_rib_property['/tw/B2'][rib_type] 

                        rp1 = (drawing_spacing+set_y-twt/2, set_z)
                        rp2 = (drawing_spacing+set_y+twt/2, set_z)
                        rp3 = (drawing_spacing+set_y-twt/2, set_z-ht)
                        rp4 = (drawing_spacing+set_y+twt/2, set_z-ht)
                        msp.add_line(rp1, rp2, dxfattribs={"layer": "Section"})
                        msp.add_line(rp1, rp3, dxfattribs={"layer": "Section"})
                        msp.add_line(rp2, rp4, dxfattribs={"layer": "Section"})
                        msp.add_line(rp3, rp4, dxfattribs={"layer": "Section"})

                        rp1 = (drawing_spacing+set_y-bt/2, set_z-ht)
                        rp2 = (drawing_spacing+set_y+bt/2, set_z-ht)
                        rp3 = (drawing_spacing+set_y-bt/2, set_z-ht-tft)
                        rp4 = (drawing_spacing+set_y+bt/2, set_z-ht-tft)
                        msp.add_line(rp1, rp2, dxfattribs={"layer": "Section"})
                        msp.add_line(rp1, rp3, dxfattribs={"layer": "Section"})
                        msp.add_line(rp2, rp4, dxfattribs={"layer": "Section"})
                        msp.add_line(rp3, rp4, dxfattribs={"layer": "Section"})

        # deal with bottom flange
        for key, item in dict_position_bot.items():
            if not pd.isna(df_section[key][run_id]):
                ### 單獨rib斷面性質提取
                rib_type = df_section[key][run_id]
                ### 提取ribs間距
                rib_spacing = re.split(r'\s*,\s*', df_section[item[0]][run_id])
                rib_num = len(rib_spacing)

                ### 參考位置定位與初始化
                rib_z_level = item[1]
                rib_y_level = item[2]
                rib_dist_y = 0
                rib_area_top_temp = 0
                ### 加勁版繪製
                for rr in range(len(rib_spacing)):
                    #### rib座標位置定出
                    set_z = -rib_z_level
                    rib_dist_y = rib_dist_y + float(rib_spacing[rr])
                    set_y = rib_y_level + rib_dist_y

                    if df_rib_property['Type'][rib_type] == 'Flat':
                        rH = df_rib_property['H/H/H'][rib_type]
                        rB = df_rib_property['B/B/B1'][rib_type]

                        rp1 = (drawing_spacing+set_y-rB/2, set_z)
                        rp2 = (drawing_spacing+set_y+rB/2, set_z)
                        rp3 = (drawing_spacing+set_y-rB/2, set_z+rH)
                        rp4 = (drawing_spacing+set_y+rB/2, set_z+rH)
                        msp.add_line(rp1, rp2, dxfattribs={"layer": "Section"})
                        msp.add_line(rp1, rp3, dxfattribs={"layer": "Section"})
                        msp.add_line(rp2, rp4, dxfattribs={"layer": "Section"})
                        msp.add_line(rp3, rp4, dxfattribs={"layer": "Section"})

                    elif df_rib_property['Type'][rib_type] == 'Tee':
                        bt = df_rib_property['B/B/B1'][rib_type] 
                        tft = df_rib_property['/tf/t'][rib_type] 
                        ht = df_rib_property['H/H/H'][rib_type] 
                        twt = df_rib_property['/tw/B2'][rib_type] 

                        rp1 = (drawing_spacing+set_y-twt/2, set_z)
                        rp2 = (drawing_spacing+set_y+twt/2, set_z)
                        rp3 = (drawing_spacing+set_y-twt/2, set_z+ht)
                        rp4 = (drawing_spacing+set_y+twt/2, set_z+ht)
                        msp.add_line(rp1, rp2, dxfattribs={"layer": "Section"})
                        msp.add_line(rp1, rp3, dxfattribs={"layer": "Section"})
                        msp.add_line(rp2, rp4, dxfattribs={"layer": "Section"})
                        msp.add_line(rp3, rp4, dxfattribs={"layer": "Section"})

                        rp1 = (drawing_spacing+set_y-bt/2, set_z+ht)
                        rp2 = (drawing_spacing+set_y+bt/2, set_z+ht)
                        rp3 = (drawing_spacing+set_y-bt/2, set_z+ht+tft)
                        rp4 = (drawing_spacing+set_y+bt/2, set_z+ht+tft)
                        msp.add_line(rp1, rp2, dxfattribs={"layer": "Section"})
                        msp.add_line(rp1, rp3, dxfattribs={"layer": "Section"})
                        msp.add_line(rp2, rp4, dxfattribs={"layer": "Section"})
                        msp.add_line(rp3, rp4, dxfattribs={"layer": "Section"})

        # 插入文字
        textp1 = drawing_spacing + (df_section['B1'][run_id]+df_section['B2'][run_id]+df_section['B3'][run_id])/2
        textp2 = -df_section['t1'][run_id]-H-df_section['t2'][run_id] -1000
        msp.add_text(
            df_section['Name'][run_id],
            dxfattribs={
                "style": text_style,  # 設定字型樣式
                "height": 300,        # 設定字高
            }
        ).set_placement((textp1, textp2), align=TextEntityAlignment.MIDDLE_CENTER)  # 設定位置

        drawing_spacing = drawing_spacing + df_section['B1'][run_id] + df_section['B2'][run_id] + df_section['B3'][run_id] + 1000

    print('> 儲存圖檔中')
    output_dxf = os.path.join(outputpath, inputfilename+"_section.dxf")
    doc.saveas(output_dxf)
    print('> 完成dxf繪製。')


class EmittingStr(QObject):
    #將stdout轉到textbrowser
    textWritten = Signal(str) 
    def write(self, text):
        self.textWritten.emit(str(text))
        loop = QEventLoop()
        QTimer.singleShot(1, loop.quit)
        loop.exec()
        QApplication.processEvents()
        
    def flush(self):
        #stdout默認有write及flush,所以須補flush method避免stdout控制錯誤
        pass

class workermct(QObject):
    finished = Signal()
    
    def __init__(self):
        QObject.__init__(self)
        
        
    def pathparameter(self,pathinput):
        self.pathparameters = pathinput
        
    def run(self):
        """"input解包"""
        inputexcelpath = self.pathparameters[0]            

        filecheck = os.path.isfile(inputexcelpath)   
        if filecheck == False:
            print('檔案不存在')
            self.finished.emit()
        
        """計算"""
        inputdata = inputexcelpath 
        Mctmainexe(inputdata)
        
        """傳出狀態"""
        self.finished.emit() 
        
class workerssc(QObject):
    finished = Signal()
    
    def __init__(self):
        QObject.__init__(self)
        
        
    def pathparameter(self,pathinput):
        self.pathparameters = pathinput
        
    def run(self):
        """"input解包"""
        inputexcelpath = self.pathparameters[0]            

        filecheck = os.path.isfile(inputexcelpath)   
        if filecheck == False:
            print('檔案不存在')
            self.finished.emit()
        
        """計算"""
        inputdata = inputexcelpath 
        Sectioncalculation(inputdata)
        
        """傳出狀態"""
        self.finished.emit() 


class workeresp(QObject):
    finished = Signal()
    
    def __init__(self):
        QObject.__init__(self)        
        
    def pathparameter(self,pathinput):
        self.pathparameters = pathinput
        
    def run(self):
        """"input解包"""
        inputexcelpath = self.pathparameters[0]            

        filecheck = os.path.isfile(inputexcelpath)   
        if filecheck == False:
            print('檔案不存在')
            self.finished.emit()
        
        """計算"""
        inputdata = inputexcelpath 
        Equivalentwidth(inputdata)
        
        """傳出狀態"""
        self.finished.emit() 


class workerefs(QObject):
    finished = Signal()
    
    def __init__(self):
        QObject.__init__(self)        
        
    def pathparameter(self,pathinput):
        self.pathparameters = pathinput
        
    def run(self):
        """"input解包"""
        inputexcelpath = self.pathparameters[0]            

        filecheck = os.path.isfile(inputexcelpath)   
        if filecheck == False:
            print('檔案不存在')
            self.finished.emit()
        
        """計算"""
        inputdata = inputexcelpath 
        Effectivesection(inputdata)
        
        """傳出狀態"""
        self.finished.emit() 


class workeras(QObject):
    finished = Signal()
    
    def __init__(self):
        QObject.__init__(self)        
        
    def pathparameter(self,pathinput):
        self.pathparameters = pathinput
        
    def run(self):
        """"input解包"""
        inputexcelpath = self.pathparameters[0]            

        filecheck = os.path.isfile(inputexcelpath)   
        if filecheck == False:
            print('檔案不存在')
            self.finished.emit()
        
        """計算"""
        inputdata = inputexcelpath 
        Allowablestress(inputdata)
        
        """傳出狀態"""
        self.finished.emit()


class workerallinone(QObject):
    finished = Signal()
    
    def __init__(self):
        QObject.__init__(self)        
        
    def pathparameter(self,pathinput):
        self.pathparameters = pathinput
        
    def run(self):
        """"input解包"""
        inputexcelpath = self.pathparameters[0]            

        filecheck = os.path.isfile(inputexcelpath)   
        if filecheck == False:
            print('檔案不存在')
            self.finished.emit()
        
        """計算"""
        inputdata = inputexcelpath 
        Sectioncalculation(inputdata)
        Effectivesection(inputdata)
        Allowablestress(inputdata)
        
        """傳出狀態"""
        self.finished.emit()


class workerplotdxf(QObject):
    finished = Signal()
    
    def __init__(self):
        QObject.__init__(self)        
        
    def pathparameter(self,pathinput):
        self.pathparameters = pathinput
        
    def run(self):
        """"input解包"""
        inputexcelpath = self.pathparameters[0]            

        filecheck = os.path.isfile(inputexcelpath)   
        if filecheck == False:
            print('檔案不存在')
            self.finished.emit()
        
        """計算"""
        inputdata = inputexcelpath 
        Plotsectiondxf(inputdata)
        
        """傳出狀態"""
        self.finished.emit()


class MainWindow(QObject):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__()
        self._window = None        
        self.setup_ui()   
        
        #將stdout轉到textbrowser
        sys.stdout = EmittingStr()
        sys.stdout.textWritten.connect(self.outputWritten) 

    @property
    def window(self):
        """The main window object"""
        self._window.setWindowTitle("SteelDeck Section Calculator")
        self._window.setWindowIcon(QIcon("./media/beam.png"))

        return self._window
    
    def setup_ui(self):
        loader = QUiLoader()
        file = QFile('./dsec.ui')
        file.open(QFile.ReadOnly)
        self._window = loader.load(file)
        file.close()
        
        self.set_button() 
        
    def outputWritten(self, text):
        """將原print到stdout內容輸出至textbrowser"""
        cursor = self._window.status.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertText(text)
        self._window.status.setTextCursor(cursor)
        self._window.status.ensureCursorVisible()
        
    def set_button(self):        
        """Setup buttons"""  
        """Choose input XML file path"""
        self._window.pushButton.clicked.connect(self.chooseexcelfilepath) 
        
        """MCT execute"""
        self._window.pushButton_2.clicked.connect(self.runmct)
        
        """SectionCalculate execute"""
        self._window.pushButton_3.clicked.connect(self.runseccal)
        
        """Open folder"""
        self._window.pushButton_4.clicked.connect(self.runof)

        """Equivalent span execute"""
        self._window.pushButton_5.clicked.connect(self.runeqspan)

        """Effective section execute"""
        self._window.pushButton_6.clicked.connect(self.runeffsec)

        """Allowable stress execute"""
        self._window.pushButton_7.clicked.connect(self.runas)

        """All in one execute"""
        self._window.pushButton_8.clicked.connect(self.runallinone)

        """Plot dxf execute"""
        self._window.pushButton_9.clicked.connect(self.runplotdxf)

    def runmct(self):
        self._window.status.setText("$ Start Execution\n")

        # Step 2: Create a QThread object
        self.mct_thread = QThread()
        # Step 3: Create a worker object
        self.mct_worker = workermct()
        # Step 4: Move worker to the thread
        self.mct_worker.moveToThread(self.mct_thread)
        # Step 5: Connect signals and slots
        self.mct_thread.started.connect(self.mct_worker.run)
        self.mct_worker.finished.connect(self.mct_thread.quit)
        self.mct_worker.finished.connect(self.mct_worker.deleteLater)
        self.mct_thread.finished.connect(self.mct_thread.deleteLater)   
        # Step 6: Set input
        'Input Parameters'
        input_excelpath = self._window.lineEdit.text()

        'Input list'
        inputparameters = [input_excelpath]
        
        '傳入worker'    
        self.mct_worker.pathparameter(inputparameters)
        # Step 7: Start the thread
        self.mct_thread.start()
        # Final resets
        self._window.pushButton_2.setEnabled(False)
        
        self.mct_thread.finished.connect(
            lambda: self._window.pushButton_2.setEnabled(True)
        )
        self.mct_thread.finished.connect(
            lambda: self._window.status.append("$ Finish Execution")
        )

    def runseccal(self):
        self._window.status.setText("$ Start Execution\n")

        # Step 2: Create a QThread object
        self.ssc_thread = QThread()
        # Step 3: Create a worker object
        self.ssc_worker = workerssc()
        # Step 4: Move worker to the thread
        self.ssc_worker.moveToThread(self.ssc_thread)
        # Step 5: Connect signals and slots
        self.ssc_thread.started.connect(self.ssc_worker.run)
        self.ssc_worker.finished.connect(self.ssc_thread.quit)
        self.ssc_worker.finished.connect(self.ssc_worker.deleteLater)
        self.ssc_thread.finished.connect(self.ssc_thread.deleteLater)   
        # Step 6: Set input
        'Input Parameters'
        input_excelpath = self._window.lineEdit.text()

        'Input list'
        inputparameters = [input_excelpath]
        
        '傳入worker'    
        self.ssc_worker.pathparameter(inputparameters)
        # Step 7: Start the thread
        self.ssc_thread.start()
        # Final resets
        self._window.pushButton_3.setEnabled(False)
        
        self.ssc_thread.finished.connect(
            lambda: self._window.pushButton_3.setEnabled(True)
        )
        self.ssc_thread.finished.connect(
            lambda: self._window.status.append("$ Finish Execution")
        )

    def runof(self):
        input_excelpath = self._window.lineEdit.text()
        (outputpath, filename_temp) = os.path.split(input_excelpath)
        try:
            os.startfile(outputpath)
        except:
            os.startfile(os.path.dirname(__file__)+'/example')
    
    def runeqspan(self):
        self._window.status.setText("$ Start Execution\n")

        # Step 2: Create a QThread object
        self.esp_thread = QThread()
        # Step 3: Create a worker object
        self.esp_worker = workeresp()
        # Step 4: Move worker to the thread
        self.esp_worker.moveToThread(self.esp_thread)
        # Step 5: Connect signals and slots
        self.esp_thread.started.connect(self.esp_worker.run)
        self.esp_worker.finished.connect(self.esp_thread.quit)
        self.esp_worker.finished.connect(self.esp_worker.deleteLater)
        self.esp_thread.finished.connect(self.esp_thread.deleteLater)   
        # Step 6: Set input
        'Input Parameters'
        input_excelpath = self._window.lineEdit.text()

        'Input list'
        inputparameters = [input_excelpath]
        
        '傳入worker'    
        self.esp_worker.pathparameter(inputparameters)
        # Step 7: Start the thread
        self.esp_thread.start()
        # Final resets
        self._window.pushButton_5.setEnabled(False)
        
        self.esp_thread.finished.connect(
            lambda: self._window.pushButton_5.setEnabled(True)
        )
        self.esp_thread.finished.connect(
            lambda: self._window.status.append("$ Finish Execution")
        )

    def runeffsec(self):
        self._window.status.setText("$ Start Execution\n")

        # Step 2: Create a QThread object
        self.efs_thread = QThread()
        # Step 3: Create a worker object
        self.efs_worker = workerefs()
        # Step 4: Move worker to the thread
        self.efs_worker.moveToThread(self.efs_thread)
        # Step 5: Connect signals and slots
        self.efs_thread.started.connect(self.efs_worker.run)
        self.efs_worker.finished.connect(self.efs_thread.quit)
        self.efs_worker.finished.connect(self.efs_worker.deleteLater)
        self.efs_thread.finished.connect(self.efs_thread.deleteLater)   
        # Step 6: Set input
        'Input Parameters'
        input_excelpath = self._window.lineEdit.text()

        'Input list'
        inputparameters = [input_excelpath]
        
        '傳入worker'    
        self.efs_worker.pathparameter(inputparameters)
        # Step 7: Start the thread
        self.efs_thread.start()
        # Final resets
        self._window.pushButton_6.setEnabled(False)
        
        self.efs_thread.finished.connect(
            lambda: self._window.pushButton_6.setEnabled(True)
        )
        self.efs_thread.finished.connect(
            lambda: self._window.status.append("$ Finish Execution")
        )

    def runas(self):
        self._window.status.setText("$ Start Execution\n")

        # Step 2: Create a QThread object
        self.as_thread = QThread()
        # Step 3: Create a worker object
        self.as_worker = workeras()
        # Step 4: Move worker to the thread
        self.as_worker.moveToThread(self.as_thread)
        # Step 5: Connect signals and slots
        self.as_thread.started.connect(self.as_worker.run)
        self.as_worker.finished.connect(self.as_thread.quit)
        self.as_worker.finished.connect(self.as_worker.deleteLater)
        self.as_thread.finished.connect(self.as_thread.deleteLater)   
        # Step 6: Set input
        'Input Parameters'
        input_excelpath = self._window.lineEdit.text()

        'Input list'
        inputparameters = [input_excelpath]
        
        '傳入worker'    
        self.as_worker.pathparameter(inputparameters)
        # Step 7: Start the thread
        self.as_thread.start()
        # Final resets
        self._window.pushButton_7.setEnabled(False)
        
        self.as_thread.finished.connect(
            lambda: self._window.pushButton_7.setEnabled(True)
        )
        self.as_thread.finished.connect(
            lambda: self._window.status.append("$ Finish Execution")
        )

    def runallinone(self):
        self._window.status.setText("$ Start Execution\n")

        # Step 2: Create a QThread object
        self.allinone_thread = QThread()
        # Step 3: Create a worker object
        self.allinone_worker = workerallinone()
        # Step 4: Move worker to the thread
        self.allinone_worker.moveToThread(self.allinone_thread)
        # Step 5: Connect signals and slots
        self.allinone_thread.started.connect(self.allinone_worker.run)
        self.allinone_worker.finished.connect(self.allinone_thread.quit)
        self.allinone_worker.finished.connect(self.allinone_worker.deleteLater)
        self.allinone_thread.finished.connect(self.allinone_thread.deleteLater)   
        # Step 6: Set input
        'Input Parameters'
        input_excelpath = self._window.lineEdit.text()

        'Input list'
        inputparameters = [input_excelpath]
        
        '傳入worker'    
        self.allinone_worker.pathparameter(inputparameters)
        # Step 7: Start the thread
        self.allinone_thread.start()
        # Final resets
        self._window.pushButton_8.setEnabled(False)
        
        self.allinone_thread.finished.connect(
            lambda: self._window.pushButton_8.setEnabled(True)
        )
        self.allinone_thread.finished.connect(
            lambda: self._window.status.append("$ Finish Execution")
        )

    def runplotdxf(self):
        self._window.status.setText("$ Start Execution\n")

        # Step 2: Create a QThread object
        self.plotdxf_thread = QThread()
        # Step 3: Create a worker object
        self.plotdxf_worker = workerplotdxf()
        # Step 4: Move worker to the thread
        self.plotdxf_worker.moveToThread(self.plotdxf_thread)
        # Step 5: Connect signals and slots
        self.plotdxf_thread.started.connect(self.plotdxf_worker.run)
        self.plotdxf_worker.finished.connect(self.plotdxf_thread.quit)
        self.plotdxf_worker.finished.connect(self.plotdxf_worker.deleteLater)
        self.plotdxf_thread.finished.connect(self.plotdxf_thread.deleteLater)   
        # Step 6: Set input
        'Input Parameters'
        input_excelpath = self._window.lineEdit.text()

        'Input list'
        inputparameters = [input_excelpath]
        
        '傳入worker'    
        self.plotdxf_worker.pathparameter(inputparameters)
        # Step 7: Start the thread
        self.plotdxf_thread.start()
        # Final resets
        self._window.pushButton_9.setEnabled(False)
        
        self.plotdxf_thread.finished.connect(
            lambda: self._window.pushButton_9.setEnabled(True)
        )
        self.plotdxf_thread.finished.connect(
            lambda: self._window.status.append("$ Finish Execution")
        )

    '''主程序執行的槽函數'''         
    @QtCore.Slot()    
    def chooseexcelfilepath(self):
        filename, filetype = QFileDialog.getOpenFileName(None, "Open file", filter='Excel (*.xlsm *.xlsx)')
        self._window.lineEdit.setText(filename)

if '__main__' == __name__:
    
    app = QApplication.instance()
    if app is None: 
        app = QApplication()
    
    mainwindow = MainWindow()
    mainwindow.window.show()

    ret = app.exec()
    sys.exit(ret)


