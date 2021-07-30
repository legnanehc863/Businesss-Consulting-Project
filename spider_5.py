# import required package
import time
from lxml import etree
from selenium import webdriver
import xlwt

# Web Scrapping using selenium
def get_html_by_selenium():
    # g2.com
    for i in range(1, 101):
        driver = webdriver.Chrome()
        driver.get(url=f'https://www.g2.com/wrike/reviews?page={i}')
        time.sleep(10)
        text = driver.page_source
        with open(r'html_data/Wrike/g2/' + str(i) + '.txt', 'w+', encoding='utf-8')as f:
            f.write(text)
        driver.quit()


# Parsing data for g2.com
def parse_1(t):
    html = etree.HTML(t)
    review = html.xpath('//div[@itemprop="review"]')
    for r in review:
        global count
        count += 1
        p_review = {}
        h = etree.tostring(r, method='html')
        h = etree.HTML(h)

        Name = ' '.join(h.xpath('//span[@itemprop="author"]//text()'))
        Position = ' '.join(h.xpath('//div[@class="c-midnight-80 line-height-h6 fw-regular"]/div/text()'))
        Company = ' '.join(h.xpath('//div[@class="c-midnight-80 line-height-h6 fw-regular"]/div/span/text()'))
        Rating = h.xpath('//div[@class="d-f mb-1"]/div/div')[0].attrib['class']
        Date = ' '.join(h.xpath('//span[@class="x-current-review-date"]/time/text()'))
        Title = ' '.join(h.xpath('//h3[@itemprop="name"]/text()'))  # 1
        Review = ' '.join(h.xpath('//div[@itemprop="reviewBody"]/div[@itemprop="reviewBody"]/div//text()'))

        p_review.setdefault('Name', Name)
        p_review.setdefault('Position', Position)
        p_review.setdefault('Company', Company)
        p_review.setdefault('Rating', Rating.split(' ')[-1])
        p_review.setdefault('Date', Date)
        p_review.setdefault('Title', Title)
        p_review.setdefault('Review', Review)
        print(count, p_review)
        all_review.setdefault(count, p_review)


def save_data_1(r):
    book = xlwt.Workbook()
    sheet = book.add_sheet('review')
    sheet.write(0, 0, 'No')
    sheet.write(0, 1, 'Name')
    sheet.write(0, 2, 'Position')
    sheet.write(0, 3, 'Company')
    sheet.write(0, 4, 'Rating')
    sheet.write(0, 5, 'Date')
    sheet.write(0, 6, 'Title')
    sheet.write(0, 7, 'Review')
    for i in range(1, len(r) + 1):
        sheet.write(i, 0, i)
        sheet.write(i, 1, r[i]['Name'])
        sheet.write(i, 2, r[i]['Position'])
        sheet.write(i, 3, r[i]['Company'])
        sheet.write(i, 4, r[i]['Rating'])
        sheet.write(i, 5, r[i]['Date'])
        sheet.write(i, 6, r[i]['Title'])
        sheet.write(i, 7, r[i]['Review'])
    book.save(r'html_data/Wrike/wrike_g2.xls')


if __name__ == "__main__":
  
    all_review = {}
    count = 0
    for i in range(1, 101):
        with open(f'html_data/Wrike/g2/{i}.txt', 'r', encoding='utf-8') as f:
            t = f.read()
            parse_1(t)
    save_data_1(all_review)
