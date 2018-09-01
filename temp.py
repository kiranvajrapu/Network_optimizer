# import pandas as pd
# import numpy as np
# ds = pd.read_excel('test_excel.xlsx')
# ds = ds.drop(['EQ', 'Frame Number', 'Event', 'EventInfo'], axis=1)
# ds = ds.loc[:, ~ds.columns.str.contains('^Unnamed')]
#
# #rename columns names at pandas
# ds = ds.rename(index=str, columns={"All-Serving Cell DL EARFCN[1]": "All-Serving Cell DL EARFCN",
#                                        "All-Serving Cell Identity[1]": "All-Serving Cell Identity"})
# ds = ds[(ds["All-Serving Cell DL EARFCN"] != 1400)]
# ds = ds[(ds['Message Type'] == 'RRC Connection Release Complete (UL-DCCH)') | (ds['Message Type'] == 'Tracking Area Update Accept')]
# ds['Time'] = ds['Time'].astype(str).str[:-3].astype(str)
# #
# # ### attach delay column
# # ds["Attach_Delay"] = np.nan
# ti = ds['Time']
#     # print(ti.values.tolist())
# tim = ti.values.tolist()
# #print(tim)
timelist = []
tim = ['11:36:40.321', '11:36:40.422', '11:36:52.545',
       '11:36:52.545','11:37:07.225', '11:37:07.326',
       '11:37:19.556', '11:37:19.556','11:37:30.653',
       '11:37:30.760', '11:37:42.406', '11:37:42.406',
       '11:37:54.129', '11:37:54.130', '11:38:04.953',
       '11:38:05.046','11:38:16.636', '11:38:16.739',
       '11:38:29.050', '11:38:29.153','11:38:37.681',
       '11:38:37.681', '11:38:37.681', '11:38:37.681',
       '11:39:05.850', '11:39:05.855', '11:39:16.568',
       '11:39:16.568','11:39:30.731', '11:39:30.834',
       '11:39:41.761', '11:39:41.761','11:39:53.671',
       '11:39:53.671', '11:40:04.485', '11:40:04.582',
       '11:40:15.896', '11:40:15.896', '11:40:26.401',
       '11:40:26.495','11:40:39.701', '11:40:39.796',
       '11:40:50.184', '11:40:50.285','11:41:02.126',
       '11:41:02.127', '11:41:13.788', '11:41:13.788',
       '10:27:10.906', '10:27:11.008', '10:27:29.104',
       '10:27:29.104','10:27:44.731', '10:27:44.731',
       '10:28:13.160', '10:28:13.278']
print(tim)
print(len(tim))
for i in range(0, len(tim) - 1):
    # print(str(tim[i+1]))
    time1 = str(tim[i])
    hours1, minutes1, seconds1 = (["0", "0", "0"] + time1.split(":"))[-3:]

    miliseconds1 = int(3600000 * int(hours1) + 60000 * int(minutes1) + 1000 * float(seconds1))

    time2 = str(tim[i + 1])

    hours2, minutes2, seconds2 = (["0", "0", "0"] + time2.split(":"))[-3:]

    miliseconds2 = int(3600000 * int(hours2) + 60000 * int(minutes2) + 1000 * float(seconds2))
        # print(miliseconds2-miliseconds1)
    hours3, milliseconds3 = divmod(miliseconds2 - miliseconds1, 3600000)
    minutes3, milliseconds3 = divmod(miliseconds2 - miliseconds1, 60000)
    seconds3 = float(miliseconds2 - miliseconds1) / 1000
    s2 = "%i:%02i:%06.3f" % (hours3, minutes3, seconds3)
    timelist.append(s2)

timelist_temp = timelist[1::2]
#timelist_temp = timelist_temp[0::2]
print(len(timelist_temp))
#print(timelist)

#
#
# #### avg
#
# avg = []
# for i in range(0, len(timelist) - 1):
#     # print(str(tim[i+1]))
#     time1 = str(timelist[i])
#     hours1, minutes1, seconds1 = (["0", "0", "0"] + time1.split(":"))[-3:]
#     miliseconds1 = int(3600000 * int(hours1) + 60000 * int(minutes1) + 1000 * float(seconds1))
#     avg.append(miliseconds1)
#
#
# print(avg)
# temp = 0
# for i in avg:
#     temp = temp + i
#
# temp = temp/len(avg)
#
#
# print(temp)
#
#
# hours3, milliseconds3 = divmod(temp, 3600000)
# minutes3, milliseconds3 = divmod(temp, 60000)
# seconds3 = float(temp) / 1000
# s2 = "%i:%02i:%06.3f" % (hours3, minutes3, seconds3)
#
#
# print(s2)

# print(tim)
print(timelist)
print(timelist_temp)
