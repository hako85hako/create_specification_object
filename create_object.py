#テキストボックス作成
def create_box_base(name,text,sheet,top,left,width,height,dot):
    s = sheet.api.Shapes.AddShape(1, top, left, width, height)
    s.Name = name
    s.TextFrame.Characters().Text = text
    s.TextFrame.HorizontalAlignment  = -4108 
    s.TextFrame.VerticalAlignment  = -4108
    if dot :
        s.Line.DashStyle = 2
    s.TextFrame.Characters().Font.Color = 1
    s.Fill.ForeColor.RGB = RGB(255, 255, 255)
    s.Line.ForeColor.RGB = RGB(0, 0, 0)

#DC3用テキストボックス作成
def create_box_DC3(name,text,sheet,top,left,width,height,dot):
    #DC3のbase作成
    DC3_base = sheet.api.Shapes.AddShape(1, top, left, width, height)
    DC3_base.Name = name
    DC3_base.TextFrame.Characters().Text = text
    DC3_base.TextFrame.HorizontalAlignment  = -4108 
    DC3_base.TextFrame.VerticalAlignment  = -4160
    if dot :
        DC3_base.Line.DashStyle = 2
    DC3_base.TextFrame.Characters().Font.Color = 1
    DC3_base.Fill.ForeColor.RGB = RGB(255, 255, 255)
    DC3_base.Line.ForeColor.RGB = RGB(0, 0, 0)

    #DC3_db作成
    DC3_db = sheet.api.Shapes.AddShape(13, top+20, left+200, width/3, height/5)
    DC3_db.Name = name
    DC3_db.TextFrame.Characters().Font.Size = 15
    DC3_db.TextFrame.Characters().Text = "DATA\r\nBASE"
    DC3_db.TextFrame.HorizontalAlignment  = -4108 
    DC3_db.TextFrame.VerticalAlignment  = -4108
    if dot :
        DC3_db.Line.DashStyle = 2
    DC3_db.TextFrame.Characters().Font.Color = 1
    DC3_db.Fill.ForeColor.RGB = RGB(157, 195, 230)
    DC3_db.Line.ForeColor.RGB = RGB(157, 195, 230)

    #DC3_soft作成
    DC3_soft = sheet.api.Shapes.AddShape(1, top+30, left+300, width/1.5, height/8)
    DC3_soft.Name = name
    #DC3_soft.TextFrame.Characters().Font.Size = 15
    DC3_soft.TextFrame.Characters().Text = "専用組み込み\r\n計測ソフトウェア"
    DC3_soft.TextFrame.HorizontalAlignment  = -4108 
    DC3_soft.TextFrame.VerticalAlignment  = -4108
    if dot :
        DC3_soft.Line.DashStyle = 2
    DC3_soft.TextFrame.Characters().Font.Color = 1
    DC3_soft.Fill.ForeColor.RGB = RGB(157, 195, 230)
    DC3_soft.Line.ForeColor.RGB = RGB(157, 195, 230)



#線作成
def create_line_base(sheet,start_x,start_y,stop_x,stop_y,dot):
    s = sheet.api.Shapes.AddLine(start_x, start_y, stop_x, stop_y).Line 
    if dot :
        s.DashStyle = 2
    s.ForeColor.RGB = RGB(0, 0, 0)

#center to center 線作成
def create_line_center(
    sheet,
    start_x,start_y,start_width,start_height,
    stop_x,stop_y,stop_width,stop_height,
    curve,dot):
    if curve == 0 :
        s = sheet.api.Shapes.AddLine(
            start_x + start_width/2, 
            start_y + start_height, 
            stop_x  + stop_width/2, 
            stop_y
            ).Line
        if dot :
            s.DashStyle = 2
        s.ForeColor.RGB = RGB(0, 0, 0) 
    #else if curve == 1 :


#right to left 線作成
def create_line_right_to_left(
    sheet,
    start_x,start_y,start_width,start_height,
    stop_x,stop_y,stop_width,stop_height,
    curve,dot):
    if curve == 0 :
        s = sheet.api.Shapes.AddLine(
            start_x + start_width, 
            start_y + start_height/2, 
            stop_x, 
            stop_y  + stop_height/2
            ).Line
        if dot :
            s.DashStyle = 2
        s.ForeColor.RGB = RGB(0, 0, 0) 
    elif curve == 2 :
        x_space = abs(( start_x + start_width ) - stop_x) 
        y_space = abs((start_y + start_height) - stop_y)
        #始点から中間まで
        s = sheet.api.Shapes.AddLine(
            start_x + start_width, 
            start_y + start_height/2, 
            start_x + start_width + x_space/2, 
            start_y + start_height/2
            ).Line
        #中間から中間
        ss = sheet.api.Shapes.AddLine(
            start_x + start_width + x_space/2,
            start_y + start_height/2, 
            start_x + start_width + x_space/2, 
            stop_y  + stop_height/2
            ).Line
        #中間から終点
        sss = sheet.api.Shapes.AddLine(
            start_x + start_width + x_space/2, 
            stop_y  + stop_height/2, 
            stop_x,
            stop_y  + stop_height/2
            ).Line
        if dot :
            s.DashStyle = 2
            ss.DashStyle = 2
            sss.DashStyle = 2
        s.ForeColor.RGB = RGB(0, 0, 0)
        ss.ForeColor.RGB = RGB(0, 0, 0)
        sss.ForeColor.RGB = RGB(0, 0, 0) 

def RGB(red, green, blue):
    return red + (green << 8) + (blue << 16)
