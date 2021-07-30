# import required package
import time
from lxml import etree
from selenium import webdriver
import xlwt


def get_html_by_selenium():
    # Trustradius.com
    for i in range(0, 150, 25):
        driver = webdriver.Chrome()
        driver.get(url=f' https://www.trustradius.com/products/monday/reviews?f={i}')
        time.sleep(10)
        text = driver.page_source
        with open(r'html_data/Monday/TR/' + str((i + 25) / 25) + '.txt', 'w+', encoding='utf-8')as f:
            f.write(text)
        driver.quit()


# Parsing data
def parse_2(t):
    html = etree.HTML(t)
    review = html.xpath('//article[@class="serp-result serp-review serp-layout"]')
    for r in review:
        global count
        count += 1
        p_review = {}
        h = etree.tostring(r, method='html')
        h = etree.HTML(h)

        Name = ' '.join(h.xpath('//div[@class="name"]//text()'))
        Position = ' '.join(h.xpath('//div[@class="position"]//text()'))
        Company = ' '.join(h.xpath('//span[@class="industry"]//text()'))
        Rating = ' '.join(h.xpath('//div[@class="trust-score__score"]//text()'))
        Date = ' '.join(h.xpath('//div[@class="review-date"]/text()'))
        Title = ' '.join(h.xpath('//div[@class="review-title"]/h3/a/text()'))
        Review = {}  
        Review_question = h.xpath('//div[@class="review-questions"]/div//h3/a/text()')
        Review_response = h.xpath('//div[@class="review-questions"]//div[@class="response"]')
        for n in range(0, len(Review_response)):
            res = Review_response[n].xpath('string(.)')
            Review.setdefault(Review_question[n], res)
        if 'Use Cases and Deployment Scope' not in Review.keys():
            Review.setdefault('Use Cases and Deployment Scope', '')

        p_review.setdefault('Name', Name)
        p_review.setdefault('Position', Position)
        p_review.setdefault('Company', Company)
        p_review.setdefault('Rating', Rating.split(' ')[1])
        p_review.setdefault('Date', Date)
        p_review.setdefault('Title', Title)
        p_review.setdefault('Review', Review)
        print(count, p_review)
        all_review.setdefault(count, p_review)


def save_data_2(r):
    book = xlwt.Workbook()
    sheet = book.add_sheet('review')
    sheet.write(0, 0, 'No')
    sheet.write(0, 1, 'Name')
    sheet.write(0, 2, 'Position')
    sheet.write(0, 3, 'Company')
    sheet.write(0, 4, 'Rating')
    sheet.write(0, 5, 'Date')
    sheet.write(0, 6, 'Title')
    sheet.write(0, 7, 'Review_' + 'Use Cases and Deployment Scope')
    sheet.write(0, 8, 'Review_' + 'Pros and Cons')
    sheet.write(0, 9, 'Review_' + 'Likelihood to Recommend')
    for i in range(1, len(r) + 1):
        sheet.write(i, 0, i)
        sheet.write(i, 1, r[i]['Name'])
        sheet.write(i, 2, r[i]['Position'])
        sheet.write(i, 3, r[i]['Company'])
        sheet.write(i, 4, r[i]['Rating'])
        sheet.write(i, 5, r[i]['Date'])
        sheet.write(i, 6, r[i]['Title'])
        sheet.write(i, 7, r[i]['Review']['Use Cases and Deployment Scope'])
        sheet.write(i, 8, r[i]['Review']['Pros and Cons'])
        sheet.write(i, 9, r[i]['Review']['Likelihood to Recommend'])
    book.save(r'html_data/Monday/monday_tr.xls')


if __name__ == "__main__":
    all_review = {}
    count = 0
    for i in range(1, 8):
        with open(f'html_data/Monday/TR/{i}.txt', 'r', encoding='utf-8') as f:
            t = f.read()
            parse_2(t)
    save_data_2(all_review)
