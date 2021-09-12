import pandas as pd
import matplotlib.pyplot as plt
from BaselineRemoval import BaselineRemoval
import beads
import deNoise
import numpy as np
# author by ChengYulong(Crystal)
# 1.读取数据
# 2.拼接
# 3.去基线
# 4.降噪
# 5.平滑

class DataProcessing:
    # 初始化属性
    def __init__(self, path):
        self.path = path
    # 读取数据，返回Excel所有数据，其为3维列表（6*n*2）
    def readData(self):
        # # 显示所有列
        # pd.set_option('display.max_columns', None)
        # # 显示所有行
        # pd.set_option('display.max_rows', None)
        # # 设置value的显示长度为100，默认为50
        # pd.set_option('max_colwidth', 100)
        allData = pd.read_excel(self.path, sheet_name=None)
        dataset = []
        for sheetName in allData.keys():
            data = pd.read_excel(self.path, sheet_name=sheetName)
            data = data.values.tolist()[5:]
            dataset.append(data)
        return dataset
    # 拼接数据,将读取的数据拼接成一个2维列表
    def splice(self):
        dataset = self.readData()
        points = []
        for i in range(1, len(dataset)):
            point = dataset[i][0][0]
            points.append(point)
        splicedData = []
        for i in range(len(dataset)-1):
            for j in range(len(dataset[i])):
                if i == 1:
                    dataset[i][j][1] -= 10000
                if points[i] > dataset[i][j][0]:
                    splicedData.append(dataset[i][j])
        return splicedData
    # 获取波长
    def getWavelength(self):
        splicedData = self.splice()
        wavelength = []
        for i in range(len(splicedData)):
            wavelength.append(splicedData[i][0])
        return wavelength
    # 获取能量值
    def getValue(self):
        splicedData = self.splice()
        value = []
        for i in range(len(splicedData)):
            value.append(splicedData[i][1])
        return value
    # 绘图
    def plot(self, wavelength, value):
        plt.rcParams['font.sans-serif'] = ['SimHei']
        plt.rcParams['axes.unicode_minus'] = False
        plt.xlabel("波长(nm)")
        plt.ylabel("能量值")
        plt.plot(wavelength, value)
        plt.show()
    # 通过BaselineRemoval函数去基线
    def removeBaseline(self, value, polynomial_degree):
        BaseObj = BaselineRemoval(value)
        Modpoly_output = BaseObj.ModPoly(polynomial_degree)
        return Modpoly_output
    # 通过beads函数去基线
    def removeBaselineByBeads(self, value, d, fc, r, Nit, lam0, lam1, lam2, pen, conv):
        removedData = beads.beads(value, d, fc, r, Nit, lam0, lam1, lam2, pen, conv)
        return list(removedData[1])
    # 降噪
    def deNoise(self, value):
        deNoisedData = deNoise.wavelet_noising(value)
        return deNoisedData.tolist()[0:len(deNoisedData)-1]
    # 噪声平滑处理
    def smooth(self, value, window_size):
        window = np.ones(int(window_size)) / float(window_size)
        return np.convolve(value, window, 'same').tolist()  # numpy的卷积函数

if __name__ == '__main__':
    path = "Sheet2.xlsx"
    dataProcessing = DataProcessing(path)
    splicedData = dataProcessing.splice()
    wavelength = dataProcessing.getWavelength()
    value = dataProcessing.getValue()
    dataProcessing.plot(wavelength, value)
    removedBaselineData1 = dataProcessing.removeBaseline(value, 10)
    removedBaselineData2 = dataProcessing.removeBaselineByBeads(value, 1, 0.3, 0.7, 60, 0.1, 0.01, 0.001, 'L1_v2', 5)
    dataProcessing.plot(wavelength, removedBaselineData1)
    dataProcessing.plot(wavelength, removedBaselineData2)
    deNoisedData = dataProcessing.deNoise(value)
    dataProcessing.plot(wavelength, deNoisedData)
    newData = dataProcessing.smooth(value, 30)
    dataProcessing.plot(wavelength, newData)