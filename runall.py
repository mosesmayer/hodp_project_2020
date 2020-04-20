import sys, os
import subprocess

filenames = [
    "Fall\\ 2016/fall_2016.xlsx",
    "Fall\\ 2017/fall_2017.xlsx",
    "Fall\\ 2018/fall_2018.xlsx",
    "Fall\\ 2019/fall_2019.xlsx",
    "Spring\\ 2016/spring_2016.xlsx",
    "Spring\\ 2017/spring_2017.xlsx",
    "Spring\\ 2018/spring_2018.xlsx",
    "Spring\\ 2019/spring_2019.xlsx",
]

cmd = "python3 Excel_to_Pandas.py"
for i in filenames:
    cmd += " {}".format(i);
subprocess.call(cmd, shell=True)
