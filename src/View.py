
from tkinter import Tk, Label, Frame, Button
from tkinter.messagebox import showinfo

class View:
    def __init__(self, dimension):
        self.root = Tk()
        self.root.title("엑셀 -> 한글 매크로")
        self.root.resizable(False, False)
        self.root.geometry(dimension)
        self.pathMap = {}      

    def addField(self, labelContent, buttonText, action, row):
        fieldLabel = self.__generateFieldLabel(labelContent)
        pathFrame = self.__generateFrame()
        pathLabel = self.__generatePathLabel(pathFrame)

        button = self.__generateButton(buttonText, lambda: self.__fdClickHandler(labelContent, action, pathLabel))

        self.__renderElements(fieldLabel, pathFrame, pathLabel, button, row)

    def addRunButton(self, buttonText, action, row):
        button = self.__generateButton(buttonText, lambda: self.__runClickHandler(action))
        button.grid(row = row, column = 2)

    def __runClickHandler(self, action):
        try: 
            excel1 = self.__readPath("사업자 정보 엑셀")
            excel2 = self.__readPath("통판 사업자 엑셀")
            hwp = self.__readPath("한글 상장 서식")

            action(excel1, excel2, hwp)


        except Exception as e:
            showinfo(
                title = "오류",
                message = str(e)
            )

        action(excel1, excel2, hwp)

    def __readPath(self, label):
        path = self.pathMap[label]

        if path != None and path != "":
            return self.pathMap[label]
        else:
            raise Exception(label + "을 선택하세요")

    def run(self):
        self.root.mainloop()

    def __generateFieldLabel(self, labelContent):
        return Label(
            self.root,
            text = labelContent,
            padx = 5,
            pady = 3
        )

    def __generateFrame(self):
        return Frame(
            self.root,
            bg = "black",
            padx = 1,
            pady = 1
        )

    def __generatePathLabel(self, frame):
        return Label(
            frame,
            bg = "white",
            width = 50
        )
    
    def __fdClickHandler(self, fileType, action, pathLabel):
        filePath = action()
        self.pathMap[fileType] = filePath
        pathLabel.config(text = filePath)

    def __generateButton(self, text, action):
        return Button(
            self.root,
            text = text,
            command = action,
            padx = 10
        )
        
    def __renderElements(self, fieldLabel, pathFrame, pathLabel, button, row):
        fieldLabel.grid(row = row, column = 0)
        pathFrame.grid(row = row, column = 1)
        pathLabel.pack()
        button.grid(row = row, column = 2)