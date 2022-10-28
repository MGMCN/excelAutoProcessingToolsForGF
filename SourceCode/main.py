import openpyxl

# path需要读取(是否需要从多个文件循环读取数据呢?)
book = openpyxl.load_workbook('20220919Summary.xlsx')

# 读取所有sheet name的名字
names = book.sheetnames
print(names)

# 这里需要读取sheet名
sheet = book["Sheet1"]
print(sheet)

# 获取 Target、 Sample、 Mean Cq的列坐标
columns = list(sheet.columns)
originalTargets = ('Target', 'Sample', 'Cq',)  # 'Cq Mean' ??? 还有啥取名儿啊 ???
targetsToColumnIndex = {}
for column in columns:
    for first_cell in column:
        coordinate = first_cell.coordinate
        value = sheet[coordinate].value
        if value in originalTargets:
            targetsToColumnIndex[value] = coordinate.__getitem__(0)
            break
# print(targetsToColumnIndex)

# 生成对应Target Sample Mean列的列数据并且对应存起来
targetsToColumn = {}
for target in targetsToColumnIndex:
    targetsToColumn[target] = list(sheet[targetsToColumnIndex[target]])
    # print(targetsToColumn[target])

duplicateMode = False  # 做单个sample设置为True

# 生成Sample->{Target(actin):Mean, Target(another):Mean}的映射数据字典
# sampleToTargetAndCq = {}
# cnt = 0
# for Target, Sample, Mean in zip(targetsToColumn['Target'], targetsToColumn['Sample'], targetsToColumn['Cq']):
#     tv = Target.value
#     sv = Sample.value
#     mv = Mean.value
#     if cnt == 0:  # 仅仅是为了跳过第一行而设置的flag
#         cnt += 1
#         continue
#     if sv not in sampleToTargetAndCq:
#         sampleToTargetAndCq[sv] = {tv: mv}
#     else:
#         sampleToTargetAndCq[sv][tv] = mv
# print(sampleToTargetAndCq)

# 生成 Target->Sample 的字典映射 targetToSampleAndCq['Target'] = {} Target有Actin和例如SOMT9、IOMT4、OMT38之类的
targetToSampleAndCq = {}
cnt = 0
targetSet = set()
for Target, Sample, Cq in zip(targetsToColumn['Target'], targetsToColumn['Sample'], targetsToColumn['Cq']):
    tv = Target.value
    sv = Sample.value
    cqv = Cq.value
    if tv is None:
        continue
    targetSet.add(tv)
    if cnt == 0:  # 仅仅是为了跳过第一行而设置的flag
        cnt += 1
        continue
    if tv not in targetToSampleAndCq:
        targetToSampleAndCq[tv] = {sv: cqv}
    else:
        targetToSampleAndCq[tv][sv] = cqv
# print(targetToSampleAndCq['Actin'])

# 选择一个 Sample 和其对应的 Actin 比较并输出结果
print(targetSet)

choosed = 'GmSOMT9'
choosedList = []
for Sample in targetToSampleAndCq[choosed]:
    # print(Sample)
    choosedList.append({'Sample': Sample, 'Cq': targetToSampleAndCq[choosed][Sample]})

# print(choosedList)  # 遍历字典的keys集合的时候输出的应该就是字典序所以可能没必要再sort一次，保险起见也可以再sort一次

choosedList = sorted(choosedList, key=lambda x: x['Sample'])


# print(choosedList)

# 每一组的数据处理函数
def calculate(actins, samples, target):
    ans = {}
    sampleCqMean = 0.0
    actinCqMean = 0.0
    size = len(samples)
    for s in samples:
        ans[s['Sample']] = {}  # 初始化3个样本 或者 1个样本
        ans[s['Sample']][target] = {}
        ans[s['Sample']]['Actin'] = {}
        sampleCqMean += s['Cq']
        actinCqMean += actins[s['Sample']]
    for s in samples:
        ans[s['Sample']][target]['CqMean'] = sampleCqMean / size
        ans[s['Sample']]['Actin']['CqMean'] = actinCqMean / size
    for s in samples:
        ans[s['Sample']]['△Cq'] = ans[s['Sample']]['Actin']['CqMean'] - ans[s['Sample']][target]['CqMean']
    for s in samples:
        ans[s['Sample']]['2△Cq'] = 2 ** ans[s['Sample']]['△Cq']
    _2deltaCqMean = 0
    for s in samples:
        _2deltaCqMean += ans[s['Sample']]['2△Cq']
    for s in samples:
        ans[s['Sample']]['2△CqMean'] = _2deltaCqMean / size
    return ans


# 当duplicateMode为True的时候pace=1，此处因为给的xlsx为单个样本所以设置为pace=1
pace = 1
cnt = 0
calculateDataSet = []
ans = None
print('choosed ->', choosed)
for Sample in choosedList:
    # print(Sample)
    # print(calculateDataSet, len(calculateDataSet))
    if cnt < pace:
        calculateDataSet.append(Sample)
        cnt += 1
    else:
        # 计算一组数据
        ans = calculate(actins=targetToSampleAndCq['Actin'], samples=calculateDataSet, target=choosed)
        print(ans)
        # 计算完了后添加这次遍历到的Sample
        calculateDataSet = []
        cnt = 1
        calculateDataSet.append(Sample)

# 这儿还得处理最后一次，因为最后一次for loop就终止了
# print(calculateDataSet, len(calculateDataSet))
ans = calculate(actins=targetToSampleAndCq['Actin'], samples=calculateDataSet, target=choosed)
print(ans)
