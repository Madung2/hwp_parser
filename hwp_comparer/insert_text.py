from time import time
import win32com.client as win32


hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
hwp.XHwpWindows.Item(0).Visible = True


before = time()
for i in range(1000):
    hwp.CreateField(Direction=f"{i}", memo=f"{i}", name=f"{i}")
    hwp.Run("MoveLineEnd")
    hwp.Run("BreakPara")
after = time()
print(f"문서 생성에 소요된 시간은 {after-before:0.2f}초입니다.")
#  약 10초 소요(최소화해놓으면 2초 소요)