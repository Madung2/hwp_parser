from pyhwpx import Hwp

# HWP 파일 열기
hwp = Hwp()
hwp.open('test1.hwp')
hwp.insert_text('Hello World')
hwp.create_table(5,5, treat_as_char=True)
for i in range(25):
    hwp.insert_text(i)
    hwp.TableRightCell()  # 기존: hwp.HAction.Run("TableRightCell")
hwp.save_as('./new.hwp')
hwp.quit()