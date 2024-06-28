from pyhwpx import Hwp

hwp = Hwp()

hwp.create_table(5,5, treat_as_char=True)
hwp.save('test_adding.table.hwpx')

## 테이블 작성은 되고 