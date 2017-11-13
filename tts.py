import os
import argparse
import win32clipboard as wc
import win32con
from gtts import gTTS
import random
from requests.exceptions import SSLError
# from pydub import AudioSegment


def get_text():
    wc.OpenClipboard()
    d = wc.GetClipboardData(win32con.CF_TEXT)
    wc.CloseClipboard()
    return d


def format_raw(raw_data, shuffle=False, transpose=True, repeat=1):
    raw_data = raw_data.strip()
    lines = raw_data.split('\r\n')
    for idx, line in enumerate(lines):
        lines[idx] = line.split('\t')
    num_each_row = [len(k) for k in lines]
    column_num = num_each_row[0]
    a = filter(lambda x: x - column_num, num_each_row)
    assert len(a) == 0
    if transpose and not shuffle:
        lines = [[l[k] for l in lines] for k in range(column_num)]
    words_queue = reduce(lambda x, y: x + y, lines)
    if shuffle:
        random.shuffle(words_queue)
    if repeat > 1:
        words_queue = map(lambda x: x + ',' + x, words_queue)
    words_queue = reduce(lambda x, y: x + ',' + y, words_queue)
    return words_queue


def parse_argument():
    parser = argparse.ArgumentParser()
    parser.add_argument('--name', default='list', type=str)
    parser.add_argument('--slow', default=False, type=bool)
    parser.add_argument('--shuffle', default=False, type=bool)
    parser.add_argument('--transpose', default=True, type=bool)
    parser.add_argument('--begin', default=1, type=int)
    parser.add_argument('--repeat', default=2, type=int)
    return parser.parse_args()


def main():
    args = parse_argument()
    name = args.name
    shuffle = args.shuffle
    slow = args.slow
    transpose = args.transpose
    begin = args.begin
    repeat = args.repeat
    if not os.path.exists('output'):
        os.mkdir('output')
    INSTRUCTION = "Copy the text for tts and " \
                  "press Enter. Input 'e' to exit."
    while raw_input(INSTRUCTION) != 'e':
        raw_data = get_text()
        raw_data = raw_data.lower()
        words_queue = format_raw(raw_data, transpose=transpose, shuffle=shuffle, repeat=repeat)
        tts = gTTS(text=words_queue, slow=slow, lang='en')
        fname = name + '_' + str(begin).zfill(4) + '.mp3'
        pname = os.path.join('output', fname)

        while True:
            try:
                tts.save(pname)
                begin += 1
                break
            except SSLError as e:
                switch = raw_input("Network error. Input 'r' to retry. Input 's' to skip")
                if switch == 's':
                    break


if __name__ == "__main__":
    main()
