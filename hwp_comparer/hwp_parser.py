import olefile
import zlib
import struct
import re
import unicodedata
import pandas as pd
# olefile은 

class HWPExtractor(object):
    FILE_HEADER_SECTION = "FileHeader"
    HWP_SUMMARY_SECTION = "\x05HwpSummaryInformation"
    SECTION_NAME_LENGTH = len("Section")
    BODYTEXT_SECTION = "BodyText"
    HWP_TEXT_TAGS = [67]

    def __init__(self, filename):
        self._ole = self.load(filename)
        self._dirs = self._ole.listdir()

        self._valid = self.is_valid(self._dirs)
        if not self._valid:
            raise Exception("Not Valid HwpFile")
        
        self._compressed = self.is_compressed(self._ole)
        self.text = self._get_text()

    def load(self, filename):
        return olefile.OleFileIO(filename)

    def is_valid(self, dirs):
        if [self.FILE_HEADER_SECTION] not in dirs:
            return False
        return [self.HWP_SUMMARY_SECTION] in dirs

    def is_compressed(self, ole):
        header = self._ole.openstream("FileHeader")
        header_data = header.read()
        return (header_data[36] & 1) == 1

    def get_body_sections(self, dirs):
        m = []
        for d in dirs:
            if d[0] == self.BODYTEXT_SECTION:
                m.append(int(d[1][self.SECTION_NAME_LENGTH:]))
        return ["BodyText/Section" + str(x) for x in sorted(m)]

    def get_text(self):
        return self.text

    def _get_text(self):
        sections = self.get_body_sections(self._dirs)
        text = ""
        for i, section in enumerate(sections):
            section_text = self.get_text_from_section(section)
            text += section_text
            text += "\n"
        self.text = text
        return self.text

    def get_text_from_section(self, section):
        bodytext = self._ole.openstream(section)
        data = bodytext.read()

        unpacked_data = zlib.decompress(data, -15) if self.is_compressed else data
        size = len(unpacked_data)

        i = 0
        idx = 0
        text = ""
        while i < size:
            header = struct.unpack_from("<I", unpacked_data, i)[0]
            rec_type = header & 0x3ff
            level = (header >> 10) & 0x3ff
            rec_len = (header >> 20) & 0xfff
            
            if rec_type in self.HWP_TEXT_TAGS:
                
                rec_data = unpacked_data[i + 4:i + 4 + rec_len]
                decode_text = rec_data.decode('utf-16')
                res = self.remove_control_characters(self.remove_chinese_characters(decode_text))
                print(f"{idx}: {res}")
                input()
                text += res
                text += "\n"
                idx+=1
            # elif rec_type in self.HWP_TABLE_TAGS:
            #     rec_data = unpacked_data[i + 4:i + 4 + rec_len]
            #     decode_text = rec_data.decode('utf-16')
            #     cleaned_text = self.remove_control_characters(self.remove_chinese_characters(decode_text))
            #     print('cleaned_text:', cleaned_text)

            i += 4 + rec_len

        return text


    def is_table_text(self, rec_type):
        # Define logic to differentiate table text from regular text if needed
        return rec_type in self.HWP_TEXT_TAGS

    def format_table(self, table):
        formatted_text = ""
        for row in table:
            formatted_text += ", ".join(row.split()) + "\n"
        return formatted_text

    @staticmethod
    def remove_chinese_characters(s: str):
        return re.sub(r'[\u4e00-\u9fff]+', '', s)

    @staticmethod
    def remove_control_characters(s):
        return "".join(ch for ch in s if unicodedata.category(ch)[0] != "C")


# Main code to run the HWPExtractor
if __name__ == "__main__":
    filename = "test1.hwp"  # Replace with your HWP file path
    extractor = HWPExtractor(filename)
    extracted_text = extractor.get_text()
