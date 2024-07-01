# import os
################################################################################# hwp5proc 작동 안됨..
# # HWP 파일 경로와 출력 XML 파일 경로 지정
# hwp_file = "/path/to/your/file.hwp"
# output_xml_path = "/path/to/your/output.xml"

# # hwp5proc 명령어 실행
# command = f'hwp5proc xml "{hwp_file}" > "{output_xml_path}"'
# os.system(command)

# # XML 파일 읽기
# # with open(output_xml_path, 'r', encoding='utf-8') as file:
# #     xml_data = file.read()
# #     print(xml_data)


import hwp5

# 새로운 hwp 파일 생성
hwp = hwp5.HWP5File()

# 섹션 생성
section = hwp.body_text.add_section()

# 제목 추가
title = section.add_paragraph()
title.add_run().text = "Introduction of Spire.Doc for Python"
title_style = title.add_run()
title_style.style_id = 1  # 기본 스타일 번호 (Heading1 같은 스타일을 정의하는 부분이 없기 때문에 기본 스타일 사용)

# 첫 번째 단락 추가
paragraph_1 = section.add_paragraph()
paragraph_1.add_run().text = (
    "Spire.Doc for Python is a professional Python library designed for developers to "
    "create, read, write, convert, compare and print Word documents in any Python application "
    "with fast and high-quality performance."
)

# 두 번째 단락 추가
paragraph_2 = section.add_paragraph()
paragraph_2.add_run().text = (
    "As an independent Word Python API, Spire.Doc for Python doesn't need Microsoft Word to "
    "be installed on neither the development nor target systems. However, it can incorporate Microsoft Word "
    "document creation capabilities into any developers' Python applications."
)

# 파일 저장
with open("WordDocument.hwp", "wb") as f:
    hwp.save(f)
