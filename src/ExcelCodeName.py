class ExcelCodeName:

    def __init__(self, code: str, columnLen: int = 1):
        self.code = code
        self.columnLen = columnLen

    def getCode(self) -> str:
        return self.code

    def nextColumn(self):
        if self.code[self.columnLen - 1] != 'Z':
            self.code = self.code[:self.columnLen - 1] + chr(
                ord(self.code[self.columnLen - 1]) + 1) + self.code[self.columnLen:]
            # print(self.code, self.columnLen)
        elif len(self.code[:self.columnLen - 1]) == 0:
            self.code = "AA" + self.code[self.columnLen:]
            self.columnLen += 1
            # print(self.code, self.columnLen)
        else:
            self.code = chr(ord(self.code[self.columnLen - 2]) + 1) + \
                "A" + self.code[self.columnLen:]
            self.columnLen += 1
            # print(self.code, self.columnLen)

    def nextRow(self):
        self.code = self.code[:self.columnLen] + \
            str(int(self.code[self.columnLen:]) + 1)

    def offset(self, columnOffset: int, rowOffset: int):
        return chr(ord(self.code[0]) + columnOffset) + str(int(self.code[1:]) + rowOffset)

    def __str__(self):
        return self.code
