import numpy as np
import xlwings as xw

import create_object as co
def app():
    wb = xw.books.active
    sheet = wb.sheets.active
    line_lists = []
    target = '_'
    # start_unit,stop_unit,line_kind,line_curve,dot
    # line_kind = 
    #         0 ：center to center
    #         1 ：right to left
    #         2 ：right to right  
    connect_list = [
        ["_ups","power_unit",2,2,False],
        ["cpu_unit","switch_hub",1,2,False],
        ["switch_hub","pcs_other",1,2,True],
        ["pcs_other","Huawei1_pcs",1,2,True],
        ["pcs_other","Huawei2_pcs",1,2,True],
        ["m_router","switch_hub",0,0,False],
        ["_DC3","switch_hub",1,2,False]
    ]

    for s in sheet.shapes: 
        s_name = str(s.name)
        for connect in connect_list:
            if connect[0] == s_name:
                for ss in sheet.shapes:
                    ss_name = str(ss.name)
                    if connect[1] == ss_name:
                        line_list = [
                            s_name +" to "+ ss_name, 
                            s.left , s.top , s.width , s.height, 
                            ss.left, ss.top, ss.width, ss.height,
                            connect[2],connect[3],connect[4]
                        ]
                        line_lists.append(line_list)

    for line in line_lists:
        print(line)
        start_x         = line[1]
        start_y         = line[2]
        start_width     = line[3]
        start_height    = line[4]
        stop_x          = line[5]
        stop_y          = line[6]
        stop_width      = line[7]
        stop_height     = line[8]
        line_kind       = line[9]
        line_curve      = line[10]
        line_dot        = line[11]

        #center to center
        if line_kind == 0 :
            co.create_line_center(
                sheet,
                start_x,start_y,start_width,start_height,
                stop_x,stop_y,stop_width,stop_height,
                curve = line_curve,
                dot = line_dot
            )
        if line_kind == 1 :
            co.create_line_right_to_left(
                sheet,
                start_x,start_y,start_width,start_height,
                stop_x,stop_y,stop_width,stop_height,
                curve = line_curve,
                dot = line_dot
            )
        
        

if __name__=="__main__":
    app()
