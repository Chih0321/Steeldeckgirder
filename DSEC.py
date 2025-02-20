import pandas as pd
import numpy as np
import re
import math
import os
import openpyxl

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
        area_girder, z_girder, y_girder, iyy_girder, izz_girder = Girdersection(run_id, df_section)

        '''加入加勁鈑'''
        ## 初始化
        dict_position_top = {'Top-Flange (Left-type)':['Top-Flange (Left-spacing)', df_section['t1'][run_id], df_section['Ref_top'][run_id]],
                            'Top-Flange (Center-type)':['Top-Flange (Center-spacing)', df_section['t1'][run_id], df_section['Ref_top'][run_id]+df_section['B1'][run_id]],
                            'Top-Flange (Right-type)':['Top-Flange (Right-spacing)', df_section['t1'][run_id], df_section['Ref_top'][run_id]+df_section['B1'][run_id]+df_section['B2'][run_id]],
                            }
        dict_position_bot = {'Bottom-Flange (Left-type)':['Bottom-Flange (Left-spacing)', df_section['t1'][run_id]+df_section['H'][run_id], df_section['Ref_bot'][run_id]],
                            'Bottom-Flange (Center-type)':['Bottom-Flange (Center-spacing)',  df_section['t1'][run_id]+df_section['H'][run_id], df_section['Ref_bot'][run_id]+df_section['B4'][run_id]],
                            'Bottom-Flange (Right-type)':['Bottom-Flange (Right-spacing)',  df_section['t1'][run_id]+df_section['H'][run_id], df_section['Ref_bot'][run_id]+df_section['B4'][run_id]+df_section['B5'][run_id]],
                            }
        section_area = area_girder
        section_zna = z_girder
        section_yna = y_girder
        section_iyy = iyy_girder
        section_izz = izz_girder

        ## 頂版加勁
        position_rib_top_y = []
        area_rib_top = []
        for key, item in dict_position_top.items():
            if not pd.isna(df_section[key][run_id]):
                rib_type = df_section[key][run_id]
                rib_area = df_rib_property['area'][rib_type]
                rib_z_p = df_rib_property['z_p_na'][rib_type]
                rib_z_n = df_rib_property['z_n_na'][rib_type]
                rib_y = df_rib_property['y_na'][rib_type]
                rib_iyy = df_rib_property['iyy'][rib_type]
                rib_izz = df_rib_property['izz'][rib_type]

                rib_spacing = re.split(r'\s*,\s*', df_section[item[0]][run_id])
                rib_num = len(rib_spacing)

                rib_z_level = item[1]
                rib_y_level = item[2]
                rib_dist_y = 0
                rib_area_top_temp = 0
                for rr in range(len(rib_spacing)):
                    area_pre = section_area
                    zna_pre = section_zna
                    yna_pre = section_yna

                    section_area = section_area + rib_area
                    rib_area_top_temp = rib_area_top_temp + rib_area
                    area_rib_top.append(rib_area)

                    set_z = rib_z_level + rib_z_n
                    rib_dist_y = rib_dist_y + float(rib_spacing[rr])
                    set_y = rib_y_level + rib_dist_y

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
                rib_type = df_section[key][run_id]
                rib_area = df_rib_property['area'][rib_type]
                rib_z_p = df_rib_property['z_p_na'][rib_type]
                rib_z_n = df_rib_property['z_n_na'][rib_type]
                rib_y = df_rib_property['y_na'][rib_type]
                rib_iyy = df_rib_property['iyy'][rib_type]
                rib_izz = df_rib_property['izz'][rib_type]

                rib_spacing = re.split(r'\s*,\s*', df_section[item[0]][run_id])
                rib_num = len(rib_spacing)

                rib_z_level = item[1]
                rib_y_level = item[2]
                rib_dist_y = 0
                rib_area_bot_temp = 0
                for rr in range(len(rib_spacing)):
                    area_pre = section_area
                    zna_pre = section_zna
                    yna_pre = section_yna

                    section_area = section_area + rib_area
                    rib_area_bot_temp = rib_area_bot_temp + rib_area
                    area_rib_bot.append(rib_area)

                    set_z = rib_z_level - rib_z_n
                    rib_dist_y = rib_dist_y + float(rib_spacing[rr])
                    set_y = rib_y_level + rib_dist_y

                    position_rib_bot_z = set_z
                    position_rib_bot_y.append(set_y)

                    section_zna = ((area_pre)*section_zna + rib_area*set_z)/section_area
                    section_yna = ((area_pre)*section_yna + rib_area*set_y)/section_area

                    section_iyy = section_iyy + area_pre*(section_zna - zna_pre)**2 + rib_iyy + rib_area*(section_zna - set_z)**2
                    section_izz = section_izz + area_pre*(section_yna - yna_pre)**2 + rib_izz + rib_area*(section_yna - set_y)**2                

        '''Shear & Torsion調整(僅含箱梁不含懸伸翼版)'''
        ## Shear
        # NOTE 在斜腹鈑時可能不是這樣算，但無可以驗證的例子
        shear_bt = df_section['B2'][run_id] +df_section['tw1'][run_id] +df_section['tw2'][run_id]
        shear_bb = df_section['B5'][run_id] +df_section['tw1'][run_id] +df_section['tw2'][run_id]
        shear_t1 = df_section['t1'][run_id]
        shear_t2 = df_section['t2'][run_id]
        shear_h = df_section['H'][run_id]
        shear_tw1 = df_section['tw1'][run_id]
        shear_tw2 = df_section['tw2'][run_id]

        section_asy = shear_bt*shear_t1 + shear_bb*shear_t2
        section_asz = shear_h*shear_tw1 + shear_h*shear_tw2

        ## Torison
        # 4A^2/(b/t)
        torsion_bt = df_section['B2'][run_id] +df_section['tw1'][run_id]/2 +df_section['tw2'][run_id]/2
        torsion_bb = df_section['B5'][run_id] +df_section['tw1'][run_id]/2 +df_section['tw2'][run_id]/2
        torsion_h = df_section['H'][run_id] +df_section['t1'][run_id]/2 +df_section['t2'][run_id]/2
        torsion_area = torsion_bb*torsion_h

        section_ixx = 4*(torsion_area)**2 /(torsion_bt/shear_t1 +torsion_bb/shear_t2 +torsion_h/shear_tw1 +torsion_h/shear_tw2)

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

        ## 腹版1貢獻
        # NOTE 用很粗略算法
        ### Zyy
        dist_zna_w1_1 = section_zna -df_section['t1'][run_id]
        zyy_w1_1 = (dist_zna_w1_1*shear_tw1)*abs(dist_zna_w1_1/2)
        dist_zna_w1_2 = section_zna_bot -df_section['t2'][run_id]
        zyy_w1_2 = (dist_zna_w1_2*shear_tw1)*abs(dist_zna_w1_2/2)
        zyy_w1 = zyy_w1_1 +zyy_w1_2
        ### Zzz
        # FIXME 斜腹版會不準
        dist_w1 = (df_section['Ref_top'][run_id] +df_section['B1'][run_id] + df_section['Ref_bot'][run_id] +df_section['B4'][run_id])/2 -df_section['tw1'][run_id]/2
        zzz_w1 = (shear_h*shear_tw1)*abs(section_yna - dist_w1)
        
        ## 腹版2貢獻
        # NOTE 用很粗略算法
        ### Zyy
        dist_zna_w2_1 = section_zna -df_section['t2'][run_id]
        zyy_w2_1 = (dist_zna_w2_1*shear_tw2)*abs(dist_zna_w2_1/2)
        dist_zna_w2_2 = section_zna_bot -df_section['t2'][run_id]
        zyy_w2_2 = (dist_zna_w2_2*shear_tw2)*abs(dist_zna_w2_2/2)
        zyy_w2 = zyy_w2_1 +zyy_w2_2
        ### Zzz
        # FIXME 斜腹版會不準
        dist_w2 = (df_section['Ref_top'][run_id] +df_section['B1'][run_id] +df_section['B2'][run_id] + df_section['Ref_bot'][run_id] +df_section['B4'][run_id] +df_section['B5'][run_id])/2 +df_section['tw1'][run_id]/2
        zzz_w2 = (shear_h*shear_tw2)*abs(section_yna - dist_w2)

        ## 頂版加勁版貢獻
        ### Zyy
        zyy_rib_top = rib_area_top_temp*abs(section_zna - position_rib_top_z)
        ### Zzz
        zzz_rib_top = 0
        for i in range(len(position_rib_top_y)):
            zzz_rib_top = zzz_rib_top +area_rib_top[i]*abs(section_yna -position_rib_top_y[i])

        ## 底版加勁版貢獻
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


