class line:
    start_x = 0
    start_y = 0
    stop_x = 0
    stop_y = 0
    tag = ""
    judge = False

    def test_method1(self):
        print("start_x:" + str(self.start_x))
        print("start_y:" + str(self.start_x))
        print("stop_x:" + str(self.stop_x))
        print("stop_y:" + str(self.stop_y))
    
    def judge_true(self,tag):
        self.tag =  tag
        self.judge = True
    
    def set_status(self,status):
        self.start_x = status[0]
        self.start_y = status[1]
        self.stop_x = status[2]
        self.stop_y = status[3]
