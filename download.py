import os
import youtube_dl
from pydub import AudioSegment
import xlsxwriter
import time
import sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import subprocess

MAX_VIDS = 1000000
pagesToSearch = int(sys.argv[2])

MashUpChunksSizeSeconds = 10
interval_between_slice = 60 * 5

duration = 'long'
if len(sys.argv) >= 4:
    duration = sys.argv[3]
    
ydl_opts = {
    'format': 'bestaudio/best',
    'postprocessors': [{
        'key': 'FFmpegExtractAudio',
        'preferredcodec': 'wav',
    }],
    'writedescription': 'writedescription',
}

DIR = os.getcwd() + os.sep

def runDownloadingCycle(requestString, pagesToSearch, duration):
    print(f'Running another downloading cycle. Searching for key words: "{requestString}"')
    innerVideoDir = f'{requestString.replace(" ", "_")}_videos{os.sep}'
    innerMashUpsDir = f'{requestString.replace(" ", "_")}_mashups{os.sep}'
    subprocess.run(f'mkdir {innerVideoDir} && mkdir {innerMashUpsDir}', shell=True)

    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=chrome_options)

    def getURLs(request_string, maxPages, dur):
        link = f"https://www.youtube.com/results?search_query={request_string.replace(' ', '+')}"
        if dur == 'long':
            link += '&sp=EgIYAg%253D%253D'
        elif dur == 'short':
            link += '&sp=EgIYAQ%253D%253D'
        driver.get(link)

        urls = []
        page = 0
        while page < maxPages:
            if page % 10 == 0:
                print(f'pages scrolling: №{page}')
            height = driver.execute_script("return document.documentElement.scrollHeight")
            driver.execute_script("window.scrollTo(0, " + str(height) + ");")
            time.sleep(5)
            page += 1

        links = driver.find_elements_by_xpath('//*[@id="video-title"]')
        print(f'Total links on scrolled {page} pages: {len(links)}')
        try:
            for link in links[:MAX_VIDS]:
                url = str(link.get_attribute("href"))
                """ELIMINATING UNACCESSIBLE VIDEOS AND STREAMS, CUTTING REDUNDANT URL SUFFIXES"""
                if url != 'None' and 'назад' in link.get_attribute('aria-label').lstrip(link.text):
                    urls.append(url.split('&')[0])
            print("Total collected good video URLs:", len(urls))
        except Exception as e0:
            print(e0)
        return urls

    urls = getURLs(requestString, pagesToSearch, duration)

    """DOWNLOADING AUDIOS BY LIST FROM videoIDsTxt """
    lines, errorLines = [], []
    try:
        for i in range(len(urls)):
            downloaded = os.listdir(f'{DIR}')
            try:
                with youtube_dl.YoutubeDL(ydl_opts) as ydl:
                    ydl.download([f'{urls[i]}'])
            except Exception as e1:
                errorLines.append((urls[i], str(e1)))
            finally:
                new = [i for i in os.listdir(f'{DIR}') if i not in downloaded]
                """2 FILES - NAME.WAV AND NAME.DESCRIPTION MUST'VE BEEN DOWNLOADED. PROCESSING:"""
                wavDownloaded = False
                description, filename = '', ''
                for newfile in new:
                    if newfile.endswith('.description'):
                        try:
                            with open(newfile, 'r', encoding='utf-8') as rdf:
                                for word in rdf.readlines():
                                    description += word
                        except Exception as e10:
                            errorLines.append((urls[i], 'e10: ' + str(e10)))
                        os.remove(newfile)
                    elif newfile.endswith('.wav'):
                        filename = newfile
                        wavDownloaded = True
                    else:
                        errorLines.append((urls[i], f'Strange file: {newfile}'))

                if not wavDownloaded:
                    print(f'№{i} of {len(urls)}  (max: {MAX_VIDS}): FAILED TO DOWNLOAD FROM {urls[i]} ')
                    errorLines.append((urls[i], 'NOT DOWNLOADED'))
                    continue
                else:
                    """REFORMATTING AUDIO"""
                    try:
                        tmp = f'{DIR}tmp.wav'
                        os.rename(f'{DIR}{filename}', tmp)
                        os.system(
                            f'ffmpeg -i "{tmp}" -acodec pcm_s16le -ac 1 -ar 16000 "{DIR}{innerVideoDir}{filename}"')
                        os.remove(tmp)
                    except Exception as e2:
                        errorLines.append((urls[i], str(e2)))
                    """WRITING MASH-UPS"""
                    mashUpDone = 'нет'
                    length = 0
                    try:
                        audio = AudioSegment.from_wav(f'{DIR}{innerVideoDir}{filename}')
                        length = round(audio.duration_seconds)
                        a, audio_mashup = 0, 0
                        for j in range(int(length / interval_between_slice)):
                            audio_mashup += audio[a: MashUpChunksSizeSeconds * 1000 + a]
                            a += interval_between_slice * 1000
                        audio_mashup.export(f'{DIR}{innerMashUpsDir}audio_mashup_{filename}', format="wav")
                        mashUpDone = f'{innerMashUpsDir}audio_mashup_{filename}'
                    except MemoryError:
                        continue
                    except Exception as e3:
                        print(f'Mashup failed: seems like too short or bad audio: {e3}')
                    finally:
                        lines.append((urls[i], innerVideoDir + filename, mashUpDone, description, length))
                        continue
    except Exception as e4:
        errorLines.append(('DOWNLOADING CYCLE CRASHED', str(e4)))

    """WRITING XLSX FILE WITH DATA"""
    wb = xlsxwriter.Workbook(f'{requestString.replace(" ", "_")}_Data.xlsx')
    bold = wb.add_format({'bold': True})

    sh = wb.add_worksheet('Data')
    sh.set_column(0, 0, 20)
    sh.set_column(1, 1, 20)
    sh.set_column(2, 2, 20)
    sh.set_column(3, 3, 120)
    sh.set_column(4, 4, 10)
    sh.set_column(5, 5, 10)
    sh.write(0, 0, 'Ссылка на видео YouTube', bold)
    sh.write(0, 1, 'Аудио', bold)
    sh.write(0, 2, 'Мэшап', bold)
    sh.write(0, 3, 'Описание', bold)
    sh.write(0, 4, 'Секунды', bold)
    sh.write(0, 5, 'Удалить (+)', bold)
    for i in range(len(lines)):
        sh.set_row(i + 1, 100)
        sh.write(i + 1, 0, lines[i][0], wb.add_format({'text_wrap': True}))
        sh.write_url(i + 1, 1, lines[i][1], wb.add_format({'text_wrap': True}))
        sh.write_url(i + 1, 2, lines[i][2], wb.add_format({'text_wrap': True}))
        for j in range(3, 5):
            sh.write(i + 1, j, lines[i][j], wb.add_format({'text_wrap': True}))

    er = wb.add_worksheet('Errors')
    er.set_column(0, 0, 70)
    er.set_column(1, 1, 70)
    er.write(0, 0, 'Ссылка на видео YouTube', bold)
    er.write(0, 1, 'Ошибка', bold)
    for i in range(len(errorLines)):
        for j in range(2):
            er.write(i + 1, j, errorLines[i][j])
    wb.close()


if __name__ == "__main__":
    print(f'Good day, ladies and gentlemen, starting our script. ')
    print(f'We will search through {pagesToSearch} Youtube pages.')
    print(f'Maximum amount of possible downloaded videos this time: {MAX_VIDS}.')
    if duration == 'long' or duration == 'short':
        print(f'Only {duration.upper()} videos will be searched and downloaded')
    if sys.argv[1].endswith('.txt'):
        with open(sys.argv[1], 'r') as rf:
            for line in rf.readlines():
                try:
                    runDownloadingCycle(line.rstrip('\n'), pagesToSearch, duration)
                except Exception as e:
                    print(e)
    else:
        runDownloadingCycle(sys.argv[1], pagesToSearch, duration)
