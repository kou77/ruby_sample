# coding: shift_jis

require 'win32ole'

def clmn2clms(cn)
	cs = ""
	while true do
		if cn > 26
			cs = ("A".ord + ((cn - 1) % 26)).chr + cs
			cn = (cn - 1) / 26
		else
			cs = ("A".ord + (cn - 1)).chr + cs
			break
		end
	end
	cs
end

=begin
p "clmn2clms(10): " + clmn2clms(10)
p "clmn2clms(26): " + clmn2clms(26)
p "clmn2clms(29): " + clmn2clms(29)
=end

def range_a1_string(tr, br, lc, rc)
	clmn2clms(lc) + tr.to_s + ":" + clmn2clms(rc) + br.to_s
end

=begin
p "range_a1_string(10, 10, 10, 10): " + range_a1_string(10, 10, 10, 10)
p "range_a1_string(20, 30, 50, 100): " + range_a1_string(20, 30, 50, 100)
=end

def read_sheet_data(so, tr, br, lc, rc)
#	p "so.class: " + so.class.to_s
#	p "so.Name: " + so.Name.to_s
	rs = range_a1_string(tr, br, lc, rc)
#	p "rs: " + rs
	rg = so.Range(rs)
	a1 = []
	rg.Rows.each do |row|
		a2 = []
		row.Columns.each do |cell|
#			p "cell.Value: " + cell.Value.to_s
			a2 << cell.Value.to_s
		end
		a1 << a2
	end
	a1
end
