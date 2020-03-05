from wox import Wox
import requests
import html5lib
import ctypes


# from https://stackoverflow.com/a/23285159
CF_TEXT = 1

kernel32 = ctypes.windll.kernel32
kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
user32 = ctypes.windll.user32
user32.GetClipboardData.restype = ctypes.c_void_p

def get_clipboard_text():
    user32.OpenClipboard(0)
    try:
        if user32.IsClipboardFormatAvailable(CF_TEXT):
            data = user32.GetClipboardData(CF_TEXT)
            data_locked = kernel32.GlobalLock(data)
            text = ctypes.c_char_p(data_locked)
            value = text.value
            kernel32.GlobalUnlock(data_locked)
            return value
    finally:
        user32.CloseClipboard()


def lookup_bing_dict(word: str) -> str:
    url_template = 'https://cn.bing.com/dict/search?q={keyword}'
    url = url_template.format(keyword=word)
    r = requests.get(url)
    html = r.text
    doc = html5lib.parse(html, namespaceHTMLElements=False)
    desc = doc.find('.//meta[@name="description"]').attrib['content']
    excepted_start = '必应词典为您提供{}的释义，'.format(word)
    if desc.startswith(excepted_start):
        desc = desc[len(excepted_start):]
    return desc


class BingDict(Wox):
    def query(self, query: str):
        results = []
        query = query.strip()
        if not query:
            query = get_clipboard_text().decode().strip()

        results.append({
            "Title": "BingDict Desc For \"{}\"".format(query),
            "SubTitle": lookup_bing_dict(query),
            "IcoPath":"Images/app.ico",
            "ContextData": "ctxData"
        })
        return results


if __name__ == "__main__":
    BingDict()
