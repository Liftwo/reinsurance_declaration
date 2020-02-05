import sched, time
import os

s = sched.scheduler(time.time, time.sleep)


def perform(cmd,inc):
    os.popen(cmd)
    s.enter(inc, 0, perform, (cmd, inc))


def exe(cmd, inc):
    s.enter(inc, 0, perform, (cmd, inc))
    s.run()


if __name__ == '__main__':
    print('show time after 1 day')
    path = 'C:\\Users\\3240\PycharmProjects\\季帳資料庫\\risk.py'
    exe(path, 30)
