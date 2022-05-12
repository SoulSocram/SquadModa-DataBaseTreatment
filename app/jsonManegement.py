import json
import tempfile

class jsonArq:

    def openFile(path):

        with open(path, 'r', encoding='utf-8') as arq, \
            tempfile.NamedTemporaryFile('w', delete=False) as out:
            data = json.load(arq)
            json.dump(data, out, ensure_ascii=False, indent=4, separators=(',',':'))
            return data

    