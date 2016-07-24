# coding: shift_jis

require './common'

pth = Dir.pwd + "/tes_get_sheet_data.xlsx"
p pth

app = WIN32OLE.new('Excel.Application')
book = app.Workbooks.Open(pth)

begin
	ary = read_sheet_data(book.Worksheets(1), 2, 10, 1, 2)
	ary.each_with_index do |a1, i|
		a1.each_with_index do |a2, j|
			p "ary[" + i.to_s + ", " + j.to_s + "]: " + ary[i][j]
		end
	end
rescue
	p $!.message
ensure
	book.Close
	app.Quit
#	p "bool.Close+app.Quit"
end
