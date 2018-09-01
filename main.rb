require 'open-uri'
require 'rubygems'
require 'nokogiri'
require 'write_xlsx'

def fetch_data(url)
  fields = ['lawlist_LawerName',
            'lawlist_LawerSex',
            'lawlist_Class',
            'lawlist_LawerqualNo',
            'lawlist_dtLawerqualNo',
            'lawlist_LawNo',
            'lawlist_qdzyzsj',
            'lawlist_zszkszysj']

  doc = Nokogiri::HTML(open(url), nil, "UTF-8")
  data = fields.map { |f| doc.css("span##{f}")[0].text}
  return data
end

def get_header_format(workbook)
  format = workbook.add_format # Add a format
  format.set_bold
  format.set_color('red')
  format.set_align('center')
  return format
end

def write_header(workbook, worksheet)
  names = ['姓名',
           '性别',
           '专业类型',
           '资格证号',
           '取得律师资格证时间',
           '执业证号',
           '取得律师执业证时间',
           '在深圳开始执业时间']

  format = get_header_format(workbook)

  names.each_with_index do |name, index|
    worksheet.write(0, index, name, format)
  end
end

def write_data(worksheet, data)
  data.each_with_index do |row, row_index|
    row.each_with_index do |col, col_index|
      worksheet.write(row_index + 1, col_index, col, nil)
    end
  end
end

def write_xlsx(file_name, data)
  workbook = WriteXLSX.new(file_name)
  worksheet = workbook.add_worksheet
  write_header(workbook, worksheet)
  write_data(worksheet, data)
  workbook.close
end

base_url = 'http://www.szlawyers.com/lawyer/'
urls = ['41f893c42f7f489fbeb4ec36c0d969a5',
        '4d4cb50989af48539e4c988a4161586e',
        '6505afe73afc4d61b585ce8d5583c424',
        '8682906bef13430384d0f6ce026c8d1a',
        'f7d324eb53d34d11880a8829ac8b90bd',
        'f21c7e6eef354c0190d5300e85d45298',
        'b6273b874c7c4c28bb95cfbd4b2b6cb2',
        '1cc4253345f843feb2878f2d3b2ecbd7',
        '15c68191b33649ebbaa5c27b6818569b',
        '80ac5289211f4bf2ae650965af327077',
        '0fc0fa3e2b9943eea9e056960b57c0b1',
        'da0efd24645f46bfa51de3611602b960',
        '33e80fad732f45e885869e05d6430a58',
        '21e815e56a494fabb1885aed8ee27541',
        'bfb72543846e4b4ebbba81d208100941',
        '8fb96c6a023a450fbd60cd4efe281c91',
        '9b80c45959324b2096c7b71aab2bd346']

datas = urls.map { |url| fetch_data(base_url + url) }
write_xlsx('laywers.xlsx', datas)
