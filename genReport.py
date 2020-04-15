from openpyxl import load_workbook
import pandas as pd
from openpyxl.drawing.image import Image

class genReport:

    def load(self):
        self.wb = load_workbook('test.xlsx')
        self.ws = self.wb['Template']
        #ws2 = wb.copy_worksheet(ws)
        self.config_ws = self.wb['config']
        self.df = pd.DataFrame(self.config_ws.values)

        self.img_config = self.wb['img_config']
        self.img_df = pd.DataFrame(self.img_config.values)

        print(self.df)
        print(self.img_df)

    def write(self,cellref, value):
        self.ws2[cellref] = value
        #print(cellref, value)

    def genReport(self):
        #for each worksheet, populate entries based upon config tab
        for i in range(1,len(self.df.index)):
            self.ws2 = self.wb.copy_worksheet(self.ws)

            for j in range(0,int(self.df.shape[1]),2):
                #write(1,1,df.iloc[i,1])
                self.write(self.df.iloc[i,j+1],self.df.iloc[i,j])

            #insert images
            for k in range(1,self.img_df.shape[0]):

                img1 = Image(self.img_df.iloc[k,0])
                img_loc = self.img_df.iloc[k,1]
                self.ws2.add_image(img1,img_loc)

    def save(self):
        self.wb.save('test.xlsx')

if __name__ == '__main__':
    tmp = genReport()
    tmp.load()
    tmp.genReport()
    tmp.save()

