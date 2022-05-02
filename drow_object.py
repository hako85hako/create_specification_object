import numpy as np
import xlwings as xw

import create_object as co

def app():
    # wb = xw.Book.caller()
    # wb.sheets[0].range('A1').value = 'Hello World!'
    wb = xw.books.active
    sheet = wb.sheets.active
    name_array = ["モバイルルータ","スイッチングハブ"]
    tag_array = ["m_router","switch_hub"]
    not_product_flgs = [False,False]
    base_top = 10
    base_left = 10
    box_height = 45.0
    box_width = 135.0
    DC3_box_height = box_height*8
    DC3_box_width = box_width*1.5

    DC3_index = 0
    DC3_buf = 0

    unit_index = 0
    unit_buf = 100

    pcs_index = 0
    pcs_buf = 0

    other_index = 0
    other_buf = 50

    ADIO_index = 0
    ADIO_buf = 0

    #DC3の追加
    name_array += ["小型計測端末\r\nDataCube3"]
    tag_array += ["_DC3"]
    not_product_flgs +=[False]

    #Gear-unitの追加
    name_array += ["電源ユニット","CPUユニット","イーサネット通信\r\nユニット","アナログ入力\r\nユニット(16点まで)"]
    tag_array += ["power_unit","cpu_unit","ether_unit","AI_unit"]
    not_product_flgs +=[False,False,False,False]

    #UPSの追加
    name_array += ["UPS"]
    tag_array += ["_ups"]
    not_product_flgs +=[True]

    #otherの追加
    name_array += ["スマートロガー\r\n3000A"]
    tag_array += ["pcs_other"]
    not_product_flgs +=[True]
    
    #PCSの追加（描画はotherの後にする）
    name_array += ["パワーコンディショナー\r\nHuawei製\r\n125kW×8台\r\n（SUN2000-125KTL-JPH0)" ]
    tag_array += ["Huawei1_pcs"]
    not_product_flgs +=[True]
    name_array += ["パワーコンディショナー\r\nHuawei製\r\n100kW×8台\r\n（SUN2000-100KTL-JPH0)" ]
    tag_array += ["Huawei2_pcs"]
    not_product_flgs +=[True]

    #AI,DI,DOの追加（描画はPCSの後にする）
    name_array += ["受変電設備","接点入力","接点出力" ]
    tag_array += ["_AI","_DI","_DO"]
    not_product_flgs +=[True,True,True]
   

    for name,tag,not_product_flg in zip(name_array,tag_array,not_product_flgs):
        #DC3の作成
        if tag[-4:] == "_DC3":
            top = base_top + DC3_buf + (DC3_box_height * DC3_index) 
            co.create_box_DC3(tag,name,sheet,base_left,top,DC3_box_width,DC3_box_height,not_product_flg)
            DC3_index += 1 

        #ユニットの作成
        if tag[-5:] == "_unit":
            top = base_top + unit_buf + (box_height * unit_index) + ( DC3_box_height * DC3_index )
            co.create_box_base(tag,name,sheet,base_left,top,box_width,box_height,not_product_flg)
            unit_index += 1 
        
        #ルーターの作成
        if tag[-7:]  == "_router":
            router_left = base_left + (box_width * 2)
            co.create_box_base(tag,name,sheet,router_left,base_top,(box_width/2),box_height,not_product_flg)
 
        #ハブの作成
        if tag[-4:]  == "_hub":
            hub_left = base_left + (box_width * 2)-( box_width/4 )
            hub_top = base_top + unit_buf
            co.create_box_base(tag,name,sheet,hub_left,hub_top,box_width,box_height,not_product_flg)

        #UPSの作成
        if tag[-4:]  == "_ups":
            ups_left = base_left + ( box_width / 4 )
            ups_top  = base_top + unit_buf/2  + ( DC3_box_height * DC3_index )
            co.create_box_base(tag,name,sheet,ups_left,ups_top,box_width/2,box_height/2,not_product_flg)

        #ロガー、その他の作成
        if tag[-6:]  == "_other":
            other_left = base_left + (box_width * 3.5)
            other_top  = base_top + other_buf + (box_height * other_index * 1.5 )
            co.create_box_base(tag,name,sheet,other_left,other_top,box_width,box_height,not_product_flg)
            other_index += 1
        
        #PCSの作成
        if tag[-4:] == "_pcs":
            if other_index > 0:
                other_index = 1      
            pcs_box_height = box_height*2
            pcs_left = base_left + (box_width * (3.5+(other_index)*1.5))
            pcs_top  = base_top + pcs_buf + (pcs_box_height * pcs_index * 1.2 )
            co.create_box_base(tag,name,sheet,pcs_left,pcs_top,box_width*1.5,pcs_box_height,not_product_flg)
            pcs_index += 1

        #AI,DI,DOの作成
        if tag[-3:] == "_AI" or tag[-3:] == "_DI" or tag[-3:] == "_DO":
            ADIO_top = base_top + ADIO_buf + pcs_index*(box_height*2*1.2) + ADIO_index*(box_height*1.2) 
            ADIO_left = base_left + (box_width * (3.5+(other_index)*1.75))
            co.create_box_base(tag,name,sheet,ADIO_left,ADIO_top,box_width,box_height,not_product_flg)
            ADIO_index += 1

if __name__=="__main__":
    app()