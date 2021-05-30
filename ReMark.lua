---------------------------------------------------
--
--	(C) Sinilga, 2021
--	https://github.com/sinilga/reMark
--
---------------------------------------------------
local reMark = {}


local tab = (" "):rep(4)
local DPI = 96
local twips = 1440/DPI	--[[ Число твипов в пикселе]]
local GapH = 54			--[[	Расстояние между ячейками в таблицах. Задается в твипах ]]
local CtrlWidth = 600 	--[[	Ширина окна (элемента TextBox), 
								в которое предполагается выводить результат. 
								Значение задаеься в пикселах ]]
--[[	Настройки RTF	]]


local RTF = {
	Header = [[\rtf1\ansi\ansicpg1251\deff0\nouicompat\deflang1049]],
	Text = "",
	Fonts = { [1]="Calibri", [2]="Cambria", [3]="Arial", [4]="Courier New", [5]="Times New Roman"},
	Colors = { 
		[1]=Color.Black, 		[2]=Color.new(43,87,154), 	[3]=Color.new(52,106,186), 
		[4]=Color.new(192,0,0), [5]=Color.new(240,240,240), [6]=Color.Blue, [7] = Color.Gray,},
	Styles = {
		["normal"] = { font = [[\f1\fs22\cf1]], para = [[\sa72]] },
		["h1"] = { font = [[\f2\fs28\caps\cf2\b]], para = [[\sb216\sa144]] },
		["h2"] = { font = [[\f2\fs28\cf3\b]], para = [[\sb144\sa36]] },
		["h3"] = { font = [[\f1\fs28\cf2]], para = [[\sb72]] },
		["cite"] = { font = [[\f1\fs22\cf1\i]], para = [[\li720\sb108\sa108]] },
		["pre"] = { font = [[\f4\fs22\cf1]], para = [[\li300\sa72]] },
		["grid"] = { font = [[\f4\fs20\cf1]],para=[[]]},
		["link"] = { font = [[\f1\fs22\cf6]] },
		["inline_pre"] = { font = [[\f4\fs22\cf1\highlight5]],para=[[\brdrbtw]] },
		["inline_strong"] = { font = [[\f1\fs22\cf4\b]] } ,
		["inline_light"] = { font = [[\f1\fs22\cf4\i]] },
		["inline_em"] = { font = [[\f1\fs22\cf4]] },
		["list"] = {font = [[\f1\fs22\cf1]], para = [[\fi-200\li400\ri0\tx200]] },
		["numlist"] = {font = [[\f1\fs22\cf1]], para = [[\fi-300\li400\ri0\tx300]] },
	},
	Bullet1 = [[\f1\fs22\fc1\b0 ]]..string.char(150),
	Bullet2 = [[\f1\fs22\fc1\b0 ]]..string.char(149),
}

RTF.Styles.gridH = {font = RTF.Styles.grid.font..[[\b]],para = RTF.Styles.grid.para}
RTF.Styles.table = {font = RTF.Styles.normal.font,para = ""}
RTF.Styles.tableH = {font = RTF.Styles.table.font..[[\b]],""}
RTF.Styles.alt_inline_pre = RTF.Styles.inline_pre

--[[	Объявление функций ]]

local ToRTF
local is_table
local render_lines
local render_para 
local render_table 
local render_grid 



local function get_dpi()
	require "win32"
	local param = [[HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics\AppliedDPI]]
	local val, err = win32.reg_get_value(param)
	if val then
		DPI = val
		twips = 1440/DPI
	end
end

local function SetStyle(name, sets)
	local style = RTF.Styles[name]
	
	if not style then
		return false, "Wrong style name"
	end
	local font_id, color_id, back_id
	if type(sets.FontName) == "string" then
		font_id = table.getkey(RTF.Fonts,sets.FontName)
		if not font_id then
			table.insert(RTF.Fonts,sets.FontName)
			font_id = #RTF.Fonts
		end
	end
	if sets.ForeColor then
		for i=1,#RTF.Colors do
			if RTF.Colors[i]:ToString() == sets.ForeColor:ToString() then
				color_id = i
				break
			end
		end
		if not color_id then
			table.insert(RTF.Colors,sets.ForeColor)
			color_id = #RTF.Colors
		end
	end
	if sets.BackColor then
		for i=1,#RTF.Colors do
			if RTF.Colors[i]:ToString() == sets.BackColor:ToString() then
				back_id = i
				break
			end
		end
		if not back_id then
			table.insert(RTF.Colors,sets.BackColor)
			back_id = #RTF.Colors
		end
	end
	local font = [[\f]]..(font_id or 1)
	font = font..[[\fs]]..2*(sets.FontSize or 11)
	if color_id then
		font = font..[[\cf]]..color_id
	end
	if back_id then
		font = font..[[\highlight]]..back_id
	end	
	if sets.FontBold then
		font = font..[[\b]]
	end
	if sets.FontItalic then
		font = font..[[\i]]
	end
	if sets.Underline then
		font = font..[[\u]]
	end
	if sets.TextTransform then
		local transform = function(str)
			if str:upper() == "CAPS" then
				font = font..[[\caps]]
			elseif str:upper() == "SCAPS" then	
				font = font..[[\scaps]]
			end	
		end
		local tr = set.TextTransform:trim():split(" ",0,true,transform)
	end
	local para = ""
	if sets.SpaceBefore then
		para = para..[[\sb]]..sets.SpaceBefore
	end
	if sets.SpaceAfter then
		para = para..[[\sb]]..sets.SpaceAfter
	end
	if sets.LeftIndent then
		para = para..[[\li]]..sets.LeftIndent
	end
	if sets.RightIndent then
		para = para..[[\ri]]..sets.RightIndent
	end
	if sets.FirstIndent then
		para = para..[[\fi]]..sets.FirstIndent
	end
	style.font = font
	style.para = para
end


do	

local strings, idx = {},1 
local num = 0 -- numbered list first item number
local para_style = "normal"
-------------------------------------------------------------------
--[[	get_para	
	Чтение очередного абзаца
	Результат:
		table {
			_string_ Style	- название стиля (таблицы возвращаются как "normal" и анализируются позже)
			_table_ Lines	- массив строк абзаца
			_number_ Num	- очередной номер для элемента нумерованного списка (int)
			_string_ Bullet	- тип маркера для элемента маркированного списка (char)
		}
]]

local function get_para()
	local style, lines = para_style, {}
	local bullet = nil
	
	repeat
		local str = strings[idx]
		local gapl = str:match("^%s*")
		str = str:gsub("\r$","")
		str = str:gsub("^%s+","")
		if str:match([[^"""+%s*$]]) and style == "pre" then
		--	окончание блока форматированного текста <pre>	
			idx = idx + 1
			para_style = "normal"
			return {Style = "pre", Lines = lines, Num = num, Bullet = bullet}
		elseif str:match([[^"""]]) and style ~= "pre" then
		--	начало блока форматированного текста <pre>	
			if #lines > 0 then
				para_style = "normal"
				return {Style = style, Lines = lines, Num = num, Bullet = bullet}
			end
			style = "pre"
			idx = idx + 1
		elseif style == "pre" then
		--	продолжение блока форматированного текста <pre>	
			table.insert(lines,gapl..str)
			idx = idx + 1
		elseif str:match("^%s*%+[%+%-:]+%s*$") and 
			not str:match("%+%+") and style ~= "grid" then
		--	начало описания форматированной таблицы 	
			if #lines > 0 then
				para_style = "normal"
				return {Style = style, Lines = lines, Num = num, Bullet = bullet}
			end
			style = "grid"
			table.insert(lines,gapl..str)
			idx = idx + 1
		elseif style == "grid" and str:match("^[%+!]") then
		--	продолжение описания форматированной таблицы 	
			table.insert(lines,gapl..str)
			idx = idx + 1
		elseif #str == 0 then
		--	пустая строка - окончание абзаца
			idx = idx + 1
			para_style = "normal"
			local res = {Style = style, Lines = lines, Num = num, Bullet = bullet}
			num = 0
			return res
		elseif str:match("^[!#]%s+") then
		-- заголовок первого уровня <H1>
			if #lines > 0 then
				para_style = "normal"
				return {Style = style, Lines = lines, Num = num, Bullet = bullet}
			end	
			style = "h1"
			str = str:gsub("^!%s+","")
			lines = {str}
			idx = idx + 1
			para_style = "normal"
			return {Style = style, Lines = lines, Num = num, Bullet = bullet}
		elseif str:match("^[!#][!#]%s+") then
		-- заголовок второго уровня <H2>
			if #lines > 0 then
				para_style = "normal"
				return {Style = style, Lines = lines, Num = num, Bullet = bullet}
			end	
			style = "h2"
			str = str:gsub("^!!%s+","")
			lines = {str}
			idx = idx + 1
			para_style = "normal"
			return {Style = style, Lines = lines, Num = num, Bullet = bullet}
		elseif str:match("^[!#][!#][!#]%s+") then
		-- заголовок третьего уровня <H3>
			if #lines > 0 then
				para_style = "normal"
				return {Style = style, Lines = lines, Num = num, Bullet = bullet}
			end	
			style = "h3"
			str = str:gsub("^!!!%s+","")
			lines = {str}
			idx = idx + 1
			para_style = "normal"
			return {Style = style, Lines = lines, Num = num, Bullet = bullet}
		elseif str:match("^[=%>]%s+[%S]") then
		-- начало цитаты <cite>
			if #lines > 0 then
				para_style = "normal"
				return {Style = style, Lines = lines, Num = num, Bullet = bullet}
			end
			style = "cite"
			str = str:gsub("^=%s+","")
			lines = {str}
			idx = idx + 1
		elseif str:match("^[%-%*]%s+[%S]") then
		-- элемент маркированного списка
			if #lines > 0 then
				para_style = "normal"
				return {Style = style, Lines = lines, Num = num, Bullet = bullet}
			end	
			style = "list"
			bullet = str:sub(1,1)
			str = str:gsub("^[%-%+%*]%s+","")
			lines = {str}
			idx = idx + 1
		elseif str:match("^[%d]+%.%s+") or str:match("^%+%s+") then
		-- элемент нумерованного списка
			if #lines > 0 then
				return {Style = style, Lines = lines, Num = num, Bullet = bullet}
			end	
			if str:match("^[%d]+%.%s+") then
				num = tonumber(str:match("^[%d]+"))
			else
				num = num + 1
			end	
			para_style = "numlist"
			style = "numlist"
			if str:sub(1,1) == "+" then
				str = str:gsub("^%+%s+","")
			else	
				str = str:gsub("^[%d]+%.%s+","")
			end	
			lines = {str}
			idx = idx + 1
		elseif str:match("^=+%s*$") then
		-- подчеркивание символами равенства: ...
			if #lines > 0 then
			--	... предыдущий текст - заголовок первого уровня
				idx = idx + 1
				para_style = "normal"
				return {Style = "h1", Lines = lines, Num = num, Bullet = bullet}
			else
				style = "normal"
				lines = {str}
			end
			idx = idx + 1
		elseif str:match("^%-+%s*$") then
		-- подчеркивание дефисами:...
			if #lines > 0 then
			--	... предыдущий текст - заголовок второго уровня
				idx = idx + 1
				para_style = "normal"
				return {Style = "h2", Lines = lines, Num = num, Bullet = bullet}
			else
				style = "normal"
				lines = {str}
			end	
			idx = idx + 1
		else	-- обычный текст
			if style == "table" and not str:match("!") then
				para_style = "normal"
				return {Style = style, Lines = lines, Num = num, Bullet = bullet}
			end
			table.insert(lines,str)
			idx = idx + 1
		end
		
	until idx > #strings
	if #lines > 0 then
		para_style = "normal"
		return {Style = style, Lines = lines, Num = num, Bullet = bullet}
	end	
end

function ToRTF(text)
	text = text:gsub("\t",tab)
	rtf_text = ""
	strings = text:split("\n")		
	idx = 1
	repeat
		local para = get_para()
		if para and #para.Lines > 0 then
			if para.Style == "pre" then
				rtf_text = rtf_text..(render_para(para) or "")
			else
				local tp, data = is_table(para)
				if tp == "table" then 
					rtf_text = rtf_text..(render_table(data) or "")
				elseif tp == "grid" then
					rtf_text = rtf_text..(render_grid(data) or "")
					if data.height < #para.Lines then
						for i=1,data.height do
							table.remove(para.Lines,1)
						end
						rtf_text = rtf_text..(render_para(para) or "")
					end
				else
					rtf_text = rtf_text..(render_para(para) or "")
				end	
			end	
		end	
	until idx > #strings
	return rtf_text
end

end	-- get_para

do -- render lines
--[[
	Формирование RTF-блока для массива строк
]]

local text = ""

--[[	join_lines
	Строки, заканчивающиея двумя пробелами, склеиваем со следующими
]]
local function join_lines(lines)
	local i = 1
	while i < #lines do
		if lines[i]:match("%s%s+$") then
			lines[i] = lines[i]:gsub("%s+$"," ")..lines[i+1]
			table.remove(lines,i+1)
		else	
			i = i + 1
		end
	end
	return lines
end


--------------------------------------------------------------
--[[	Работа со стилями:
	change_style:	помещает информацию о текущем используемом стиле в стек 
					и записывает в выходной блок информацию о новом стиле
	restore_style:	восстанавливает предыдущий стиль из стека				
]]

local stack = {}
local cur_style = { --[[bool b,i,u, int h]] }

local function change_style(new_style)
	if new_style == "i" and not cur_style.native_i then
		text = text.."\\i "
		cur_style.i = true
	elseif new_style == "b" and not cur_style.native_b then
		text = text.."\\b "
		cur_style.b = true
	elseif new_style == "bi" then
		if not cur_style.native_i then
			text = text.."\\i "
			cur_style.i = true
		end	
		if not cur_style.native_b then
			text = text.."\\b "
			cur_style.b = true
		end	
	else
		if cur_style then
			table.insert(stack,cur_style)
		end	
		cur_style = {name = new_style, set = RTF.Styles[new_style] or RTF.Styles["normal"]}
		for item in cur_style.set.font:gmatch("\\(.-)") do
			cur_style.native_b = item:match("^b%d*$") and item ~= "b0"
			cur_style.native_i = item:match("^i%d*$") and item ~= "i0"
			cur_style.native_u = item:match("^u%d*$") and item ~= "u0"
			cur_style.native_h = item:match("^highlight%d*$") and item ~= "highlight0"
		end
		text = text .. "\\plain{\\*\\"..cur_style.name.."}"..(cur_style.set.para or "")..(cur_style.set.font or "").." "
	end	
end

local function restore_style(style)
	if style == "i" and cur_style.i then
		text = text.."\\i0 "
		cur_style.i = false
	elseif style == "b" and cur_style.b then
		text = text.."\\b0 "
		cur_style.b = false
	elseif style == "bi" then 
		if cur_style.b then
			text = text.."\\b0 "
			cur_style.b = false
		end	
		if cur_style.i then
			text = text.."\\i0 "
			cur_style.i = false
		end	
	else
		cur_style = table.remove(stack)
		if cur_style and cur_style.name and cur_style.set then
			text = text .. "\\plain{\\*\\"..cur_style.name.."}"..(cur_style.set.para or "")..(cur_style.set.font or "").." "
		else
			text = text .. "\\plain"
		end	
	end	
end
--------------------------------------	

local function typing_rus(str, prev)
--[[
	Автозамена в соответствии с правилами набора текста:
		* прямых кавычек - парными;
		* дефисов, отбитых пробелами, - длинным тире;
		* дефисов между цифрами - коротким тире
]]
	local chset = {
		hips = string.char(150,151),
		mdash = string.char(151),
		ndash = string.char(150),
		nbsp = string.char(160),
		lquote = string.char(171),
		rquote = string.char(187)
	}

	local ch = str:sub(1,1)
	local stq, pos, out = false, 1, ""
	local prev_ch, next_ch = prev or " ",""
	repeat
		if pos == #str then
			next_ch = "."
		else	
			next_ch = str:sub(pos+1,pos+1)
		end	
		if ch == "\\" and pos < #str then
			out = out..ch..next_ch
			pos = pos + 1
			prev_ch = next_ch
			ch = str:sub(pos+1,pos+1)
		elseif ch == "\"" then
			if prev_ch:match("[%s%(]") or prev_ch:match("[%p\"%)]") and next_ch:match("%w") then
				out = out..chset.lquote
				stq = true
			elseif prev_ch:match("%w") and next_ch:match("%w") then
				out = out..(stq and chset.rquote or chset.lquote)
				stq = not stq
			else	
				out = out..chset.rquote
				stq = false
			end	
			prev_ch = ch
			ch = next_ch
		elseif ch:match("[%-"..chset.hips.."]") and prev_ch == " " and next_ch == " " then
			out = out:sub(1,-2)..chset.nbsp..chset.mdash
			prev_ch = chset.mdash
			ch = next_ch
		elseif ch:match("[%-"..chset.hips.."]") and prev_ch:match("%d") and next_ch:match("%d") then
			out = out:sub(1,-2)..chset.nbsp..chset.ndash
			prev_ch = chset.mdash
			ch = next_ch
		else	
			stq = ch == chset.lquote or stq 
			out = out..ch
			prev_ch = ch
			ch = next_ch
		end
		pos = pos + 1
	until pos > #str
	str = out
	return str,prev_ch
end

local function subst_ch(str)
--[[
	Подстановка спецсимволов HTML
]]
	local t = table.invert({"","","sbquo","","bdquo","hellip","dagger","Dagger","euro","permil","","lsaquo","","","","","","lsquo","rsquo","ldquo","rdquo","bull","ndash","mdash","","trade","","rsaquo","","","","","nbsp","","","","curren","","brvbar","sect","","copy","","laquo","not","shy","reg","","deg","plusmn","","","","micro","para","middot"})
	local function __subst(s)
		if t[s] then 
			return string.char(t[s]+127) 
		elseif s:match("^#%d+$") then	
			return string.char(tonumber(s:match("%d+")))
		else 
			return "&"..s..";"
		end	
	end	
	str = str:gsub("&(.-);", __subst)
	return str
end

local function safe_rtf(str)
--[[
	Экранирование символов разметки RTF
]]
	local RTF_mark = [[%\%{%}]]
	local out = ""
	for i=1,#str do
		local ch = str:sub(i,i)
		if ch:match("["..RTF_mark.."]") then
			out = out .. "\\" .. ch
		else
			out = out .. ch
		end
	end	
	return out
end

local lines
local idx, str, pos = 1,"",1

local function next_item(state)
	local auto = {
		["start"] = {
			["_"] = {new_state = "b1", shift = 1, skip = true},
			["*"] = {new_state = "a1", shift = 1, skip = true},
			["\""] = {new_state = "q1", shift = 1},
			["="] = {new_state = "e1", shift = 1},
			["else"] = {shift = 1},
		},
		["b1"] = {
			["_"] = {new_state = "b2", shift = 1, skip = true},
			["*"] = {new_state = "a1", shift = 1, skip = true},
			["else"] = {new_state = "i", shift = 0, skip = true, term = true},
		},	
		["b2"] = {
			["_"] = {new_state = "bi", shift = 1, skip = true, term = true},
			["*"] = {new_state = "a1", shift = 1, skip = true},
			["else"] = {new_state = "b", shift = 0, skip = true, term = true},
		},	
		["a1"] = {
			["*"] = {new_state = "a2", shift = 1, skip = true},
			["_"] = {new_state = "b1", shift = 1, skip = true},
			["else"] = {new_state = "inline_light", shift = 0, skip = true, term = true},
		},	
		["a2"] = {
			["*"] = {new_state = "inline_strong", shift = 1, skip = true, term = true},
			["_"] = {new_state = "b1", shift = 1, skip = true},
			["else"] = {new_state = "inline_em", shift = 0, skip = true, term = true},
		},	
		["q1"] = {
			["\""] = {new_state = "inline_pre", shift = 1, skip = true, term = true, clip = 1},
			["="] = {new_state = "e1", shift = 1},
			["_"] = {new_state = "b1", shift = 1, skip = true},
			["*"] = {new_state = "a1", shift = 1, skip = true},
			["else"] = {new_state = "start", shift = 1},
		},	
		["e1"] = {
			["\""] = {new_state = "q1", shift = 1},
			["="] = {new_state = "alt_inline_pre", shift = 1, skip = true, term = true, clip = 1},
			["_"] = {new_state = "b1", shift = 1, skip = true},
			["*"] = {new_state = "a1", shift = 1, skip = true},
			["else"] = {new_state = "start", shift = 1},
		},	
		["inline_pre"] = {
			["\""] = {new_state = "pre_q1", shift = 1},
			["else"] = {shift = 1},
		},
		["pre_q1"] = {
			["\""] = {new_state = "inline_pre", shift = 1, skip = true, term = true, clip = 1},
			["else"] = {new_state = "start", shift = 1},
		},	
		["alt_inline_pre"] = {
			["="] = {new_state = "pre_e1", shift = 1},
			["else"] = {shift = 1},
		},
		["pre_e1"] = {
			["="] = {new_state = "alt_inline_pre", shift = 1, skip = true, term = true, clip = 1},
			["else"] = {new_state = "start", shift = 1},
		},	
	}
	if not state then
		state = "start"
	end	
	local out = ""
	repeat
		local ch = str:sub(pos,pos)
		local next_ch = str:sub(pos+1,pos+1)
		if ch == "\\" and not state:match("pre$") and pos < #str then
			if next_ch >= "0" and next_ch <= "0" then
				local ed = str:find("%D", pos+1)
				out = out..str:sub(pos,pos + ed - 1)
				pos = pos + ed
			else
				out = out..str:sub(pos,pos+1)
				pos = pos + 2
			end	
		else
			local proc = auto[state][ch] or auto[state]["else"]
			if proc.new_state then
				state = proc.new_state
			end	
			if proc.shift then
				pos = pos + proc.shift
			end
			if not proc.skip then
				out = out..ch
			end
			if proc.clip then
				out = out:sub(1,-1*proc.clip-1)
			end
			if proc.term then
				return state, out
			end
		end	
	until pos > #str	
	return state, out
end


function render_lines(para)
	lines = join_lines(para.Lines)
	text = ""
	cur_style = {}
	stack = {}
	change_style(para.Style)
	local state = "start"
	local bold, italic = false, false
	local last_ch
	
	idx, pos = 1, 1
	str = lines[idx] or ""
	repeat
		if idx < #lines and pos > #str then
			idx = idx + 1
			text = text.."\\line "
			str = lines[idx]
			pos = 1
		elseif idx >= #lines and pos > #str then
			break
		end
		local tp, ch 
		if cur_style and cur_style.name:match("_pre$") then
			tp, ch = next_item(cur_style.name)
		else	
			tp, ch = next_item()
		end
		if cur_style and cur_style.name:match("_pre$") then
			text = text..safe_rtf(ch)
		else
			ch, last_ch = typing_rus(ch, last_ch) 
			ch = subst_ch(ch) 
			ch = ch:gsub("\\(.)","%1")
			text = text..safe_rtf(ch) 
		end
		if cur_style and tp == cur_style.name then
			restore_style()
		elseif tp == "i" then	
			if not italic then
				change_style("i")
				italic = true
			else	
				restore_style("i")
				italic = false
			end
		elseif tp == "b" then	
			if not bold then
				change_style("b")
				bold = true
			else	
				restore_style("b")
				bold = false
			end
		elseif tp == "bi" then	
			if not bold or not italic then
				change_style("b")
				bold = true
				change_style("i")
				italic = true
			else	
				restore_style("bi")
				bold = false
				italic = false
			end
		elseif tp ~= "start" then
			change_style(tp)
		end
	until pos > #str and idx >= #para.Lines
	return text
end

end	-- render lines

--[[	render_para
	Формирование RTF-блока для абзаца
]]
function render_para(para)
	local style = RTF.Styles[para.Style] or RTF.Styles["normal"]
	if not style.para then
		style.para = ""
	end
	if not style.font then
		style.font = ""
	end
	local text 
	if para.Style == "pre" then
		text = "\\plain {\\*\\pre}"..style.para..style.font.." "..table.concat(para.Lines,"\\line ")
	elseif para.Style == "normal" and #para.Lines == 1 and para.Lines[1]:match("^[%-%=]+%s*$") then
		text = style.para..style.font.." "
	else	
		text = render_lines(para)
	end	
	local block = "\\pard "
	if para.Style == "list" and para.Bullet == "-" then
		block = block..RTF.Bullet1.."\\tab "..style.para..style.font
	elseif	para.Style == "list" then
		block = block..RTF.Bullet2.."\\tab "..style.para..style.font
	elseif para.Style == "numlist" then	
		block = block..style.para..style.font.." "..para.Num..".\\tab "
	else	
		block = block..style.para..style.font
	end	
	block = block..text.."\\par"
	return block
end

do --	tables and grides
--[[	
	Обработка таблиц-списков и таблиц-сеток
	
	Структура описания таблицы:
	{	--	прим. hdrH, height, width - только для сеток	--
		_table_ columns:	массив описаний настроек столбцов
		_table_ cells:		матрица ячеек
		_number_ hdrH:		число строк заголовка (int)
		_number_ height:	число текстовых строк, относящихся к таблице (int)
		_number_ width:		общая ширина столбцов (в символах, int)
		_number_ rows:		число строк в таблице (int)
		_number_ cols:		число столбцов в таблице (int, ::= #columns)
		
	}	
	Структура описания столбца	(элемент columns)
	{	--	прим. left, right - только для сеток	--
		_number_ left: 	позиция в текстовой строке, с которой начинается столбец
		_number_ right:	позиция в текстовой строке, на которой заканчивается столбец 
						(включая разделитель)
		_string_ align: горизонтальное выравнивание содержимого столбца (ql,qr,qc)
		_number_ width:	ширина столбца в твипсах
		_number_ x:		позиция правой границы столбца в твипсах (значение для \cellx)
	}
	Структура описания ячейки (элемент cells)
	{	--	прим. rowspan, colspan, status - только для сеток	--
		_string_ text:		содержимое ячейки
		_number_ rowspan:	число объединяемых ячееек по вертикали
		_number_ colspan:	число объединяемых ячеек по горизонтали
		_string_ status:	статус ячейки (
							clvmgf - первая ячейка в объединении (левая верхняя), 
							clvmrg - входит в оъединение и находится в самом левом столбце,
									кроме самой верхней (выводится пустой)
							null - другие ячейки, входящие в объединение (не выводятся)
							normal - не входит в объединение
	}
]]

----------------------------------------------------------------

--[[ 
	Автоподбор ширины столбцов 
]]

local function create_label(text, style)
	if not style then 
		style = "normal"
	end
	local lab = Label.new("textlab",text)
	local font = RTF.Styles[style].font
	local font_num = font:match("\\f(%d+)") or "1"
	lab.FontName = RTF.Fonts[tonumber(font_num)]
	local sz = tonumber(font:match("\\fs(%d+)") or "22")
	lab.FontSize = math.ceil(sz / 2)
	lab:AdjustMinSize()
	return lab
end

local function col_width_init_data(t,style)
--[[
	cell.word_widths = {},
	cell.textlab
	
	column.minwidth
	column.maxwidth
	column.medwidth
	column.widths_data
]]
	local cells = t.cells
	local cols = t.columns
	for j=1,t.cols do
		if not cols[j].widths_data then
			cols[j].widths_data = {}
			cols[j].minwidth = 0
			cols[j].maxwidth = 0
		end
		local tmp = cols[j]
		
		for i=1,t.rows do
			local cell = cells[i][j]
			if not cell or not cell.text then
				continue
			end
			cell.textlab = create_label(cell.text, style)
			local _w = cell.textlab.Width
			local _n = cell.colspan or 1
			_w = math.ceil(_w / _n)
			for k = 1,_n do
				if not cols[j+k-1].widths_data then
					cols[j+k-1].widths_data = {}
					cols[j+k-1].maxwidth = 0
					cols[j+k-1].minwidth = 0
				end
				table.insert(cols[j+k-1].widths_data, _w)
				cols[j+k-1].maxwidth = math.max(cols[j+k-1].maxwidth, _w)
			end	
			cell.word_widths = {}
			local words = cell.text:split(" ")
			for k = 1,#words do
				local word, ww = words[k],0
				local lab = create_label(word,style)
				ww = lab.Width
				table.insert(cell.word_widths,lab.Width)
				local col_min = cols[j].minwidth
				if lab.Width > cols[j].minwidth then
					cols[j].minwidth = lab.Width
				end
				lab:DeleteControl()
			end
		end
		table.sort(cols[j].widths_data)
	end
end

local function percent_width(t, level)
	local width = 0
	local data = {}
	for j=1,t.cols do
		local w_data = t.columns[j].widths_data
		local _w = math.max(w_data[level] or w_data[1], t.columns[j].minwidth)
		data[j] = math.ceil(_w*twips)+2*GapH
		width = width + data[j]
	end
	return width, data
end

local function set_col_widths(t,style,width)
	if not width then
		width = CtrlWidth
	end
	col_width_init_data(t,style)
	
	total_width = percent_width(t,t.rows)
	local max_twips = math.floor(width*twips)
	
	if percent_width(t,1) > max_twips then
	--	wide
		for j=1,t.cols do
			t.columns[j].width = t.columns[j].minwidth
			t.columns[j].twips = math.ceil(t.columns[j].width*twips) + 2*GapH
		end
	elseif 	total_width <= max_twips then
	--	narrow 
		for j=1,t.cols do
			t.columns[j].width = t.columns[j].maxwidth
			t.columns[j].twips = math.ceil(t.columns[j].width*twips) + 2*GapH
		end
	else 
		local level, p_width, pw_data = t.rows, 0, {}
		repeat
			level = level - 1
			p_width, pw_data = percent_width(t,level)
			if p_width <= max_twips then
				break
			end
		until level < math.floor(t.rows/2)
		
		for j=1,#t.columns do
			local k = max_twips/p_width
			local a, b = 0,0
			local col = t.columns[j]
			local _min = math.ceil(col.minwidth*twips)+2*GapH
			local _max = math.ceil(col.maxwidth*twips)+2*GapH
			local new_w = math.ceil(k * pw_data[j])
			if new_w < _min then
				col.width = col.minwidth
				col.twips = _min
				max_twips = max_twips - a - col.twips
				p_width = p_width - b - pw_data[j]
			elseif	new_w > _max then
				col.width = col.maxwidth
				col.twips = _max
				max_twips = max_twips - a - col.twips
				p_width = p_width - b - pw_data[j]
			else
				col.width = new_w
				col.twips = new_w
				a = a + col.width
				b = b + pw_data[j]
			end	
		end	
	end
	local x = 0
	for j=1,#t.columns do
		x = x + t.columns[j].twips
		t.columns[j].x = x
	end
end
------------------------------------------------------

--[[
	Обработка таблиц-списков
	
	parse_table:	разбор текста, относящегося к таблице, 
					и формирование стркутуры с описанием содежимого ячеек 
					и настроек столбцов
	render_table:	формирование на основе полученной структуры RTF-блока
]]
----------------------------------------------------------------------
function render_table(t)
	local text 
	local block = ""
	local row_h = [[\trowd\trautofit1\trgaph]]..GapH
	set_col_widths(t,"table",CtrlWidth)
	for i=1,t.rows do
		local row = t.cells[i]
		local row_stru = row_h
		local row_data = ""
		for j = 1,t.cols do
			local cell = row[j]
			if not cell then
				cell = {text=""}
			end
			row_stru = row_stru..[[\clbrdrt\brdrs\brdrw10\clbrdrl\brdrs\brdrw10\clbrdrb\brdrs\brdrw10\clbrdrr\brdrs\brdrw10]]
			if i <= t.hdrH then
				row_stru = row_stru ..[[\clcbpat5\clvertalc]]
				text = render_lines({Style = "tableH", Lines = {cell.text}})
			else
				text = render_lines({Style = "table", Lines = {cell.text}})
			end
			row_stru = row_stru .. [[\cellx]]..t.columns[j].x
			row_data = row_data..[[\pard\intbl\]]
			row_data = row_data..((i <= t.hdrH and "qc") or t.columns[j].align)
			row_data = row_data.." "..text..[[\cell ]]
		end	
		block = block..row_stru..row_data..[[\row ]]
	end
	return block..[[\pard\par]]
end

local function parse_table(lines)
	local t = {}
	
	local function is_table_div(str)
		return str:match("^[%-!:]+%s*$") and
		str:match("!") and str:match("%-") and
		not str:match("([!:])%1")
	end
	
	local function table_columns(str)
		str = str:gsub("^!",""):gsub("%s*$","")
		if str:sub(-1,-1) ~= "!" then str = str.."!" end
		local columns = {}
		for item in str:gmatch("(.-)!") do
			local col = {}
			if item:match("^:%-+:$") then
				col.align = "qc"
			elseif item:match("^%-+:$") then
				col.align = "qr"
			elseif item:match("^:?%-+$") then
				col.align = "ql"
			else
				return nil
			end
			table.insert(columns,col)
		end
		return columns
	end

	t.cells = {}
	if is_table_div(lines[1]) then
		t.columns = table_columns(lines[1])
		if type(t.columns) ~= "table" then
			return nil 
		elseif type(TableAutoHrd) == "table" then	
			t.cells[1] = {}
			for i=1, #t.columns do
				t.cells[1][i] = TableAutoHrd[i]
			end
			t.hdrH = 1
		else
			t.hdrH = 0
		end
		t.cols = #t.columns
	elseif #lines > 1 and is_table_div(lines[2]) then
		t.columns = table_columns(lines[2])
		if type(t.columns) == "table" then
			t.cols = #t.columns
			t.hdrH = 1
		end	
	else	
		return nil
	end	
	
	for i=1,#lines do
		if i == t.hdrH + 1 then 
			continue
		end
		local row = {}
		local str, pos = lines[i], 1
		local cell = ""
		repeat
			local ch = str:sub(pos,pos)
			if ch == "!" then
				cell = cell:gsub("^%s+",""):gsub("%s+$","")
				table.insert(row,{text = cell})
				cell = ""
				pos = pos + 1
			elseif ch == "\\" and pos < # str then
				cell = cell..str:sub(pos, pos + 1)
				pos = pos + 2
			elseif ch == "\"" and str:sub(pos+1,pos+1) == "\"" then
				local _, ed = str:find([[.-""+]], pos + 2)
				if not ed then
					ed = #str
				end	
				cell = cell..str:sub(pos,ed)
				pos = ed + 1
			else
				cell = cell..ch
				pos = pos + 1
			end
		until pos > # str
		if cell ~= "" then
			table.insert(row,{text = cell})
		end
		table.insert(t.cells,row)
	end
	t.rows = #t.cells
	return t
end

---------------------------------------------------------------------------------
--[[
	Обработка таблиц-сеток (grid)
	is_grid_divider:	проверка, является ли переданная текстовая строка 
						разделителем рядов сетки ("+---+---:+---")
	is_grid_header:		проверка, является ли переданная текстовая строка
						разделителем заголовка сетки ("+=====+=====+")
	get_grid_columns:	получение информации о столбцах сводной таблицы
	get_grid_area:		определение размеров сетки
	set_col_widths:		подбор ширины столбцов по содержимому
	parse_grid:			разбор строк и заполнение структуры с описанием сетки
	render_grid:		формирование rtf-блока
]]

local function is_grid_divider(str)
	return str:match("^%s*[%+!][%+%-:]+%s*$") and not str:match("%+%+") and not str:match("!!")
end

local function is_grid_header(str)
	return str:match("^%s*[%+!][=%+:!]+%s*$") and not str:match("%+%+") and not str:match("!!")
end

local function get_grid_columns(str)
	local columns = {}
	local col
	local pos = str:find("+")
	if str:sub(-1,-1) ~= "+" then
		str = str .. "+"
	end
	local state = "start"
	repeat
		local ch = str:sub(pos,pos)
		if ch == "-" and state ~= "ra" then
			state = "-"
		elseif ch == ":" and state == "+" then
			state = "la"
		elseif ch == ":" and state == "-" then
			state = "ra"
		elseif ch == "+" and table.getkey({"start","ra","-"},state) then
			if type(col) == "table" then
				col.right = pos
				local s = str:sub(col.left,col.right-1)
				if s:match("^:%-+:$") then
					col.align = "qc"
				elseif s:match("^%-+:$") then
					col.align = "qr"
				else
					col.align = "ql"
				end
				table.insert(columns,col)
			end	
			col = {left = pos + 1}
			state = "+"
		else
			return nil
		end
		pos = pos + 1
	until pos > #str
	return columns
end

local function get_grid_area(lines)
--[[	область таблицы (строки из lines, относящиеся к таблице)	
		сканируем вниз вдоль левой границы, пока не встретится символ,
		отличный от "!" и "+""
]]		
	local left = lines[1]:find("%+") -- левая граница таблицы
	local w = lines[1]:find("%S+%s*$")
	local st = "+"
	local j = 2
	while j <= #lines do
		local ch = lines[j]:sub(left,left)
		local len = lines[j]:find("%S%s*$")
		if ch == "!" or ch == "+" and st == "!" then
			st = ch
			w = math.max(w,len)
		elseif j > 2 then
			break
		else	
			return nil
		end
		j = j + 1
	end
	return j-1, w
end

local function parse_grid(lines)
	local t = { }
	if not is_grid_divider(lines[1]) then
		return
	else	
		t.columns = get_grid_columns(lines[1])
		t.cols = #t.columns
	end
	t.height, t.width = get_grid_area(lines)
		
--	чтение таблицы
	local cells = {}
	local rows = {}			-- cur row indexes
	local col_index = 1
	local function max_row()
		local max = rows[1] or 0
		for j = 2,#t.columns do
			if rows[j] and rows[j] > max then 
				max = rows[j]
			end
		end
		return max
	end
	local function start_row()
		local max = max_row()
		for j = 1,#t.columns do
			if rows[j] and cells[rows[j]][j] and max > rows[j] then
				cells[rows[j]][j].rowspan = max - rows[j] + 1
			end	
			rows[j] = max + 1
		end
	end
	
	local function current_cell()
		if not cells[rows[col_index]] then
			cells[rows[col_index]] = {}
		end
		if not cells[rows[col_index]][col_index] then
			cells[rows[col_index]][col_index] = {text = ""}
		end
		return cells[rows[col_index]][col_index]
	end
	
	start_row()
	
	local idx = 2
	repeat
		local str = lines[idx]
		local div = false
		if str:match("^[%-%+%!]+%s*$") then
		--	полный разделитель строки
			if idx < t.height then
				start_row()
			end	
		elseif not t.hdrH and is_grid_header(str) then
		-- разделитель заголовка таблицы
			t.hdrH = max_row()
			start_row()
		else
			local max = max_row()
			col_index = 1
			local j = 1
			repeat
				local col = t.columns[j]
				local s = str:sub(col.left,col.right)
				if s:match("^[%+%-]+!?$") or j == #t.columns and s:match("^[%-%+]+!?%s*$") then
				--	локальный разделитель строк
					local cell = current_cell()
					if max > rows[j] then
						cell.rowspan = max - rows[j] + 1
					end	
					rows[j] = max + 1
					col_index = col_index + 1
				else
				--	обычный текст
					local cell = current_cell()
					if s:match("[!%+%-]$") or j == #t.columns then
						local cw = col.right-col.left+1
						local last_ch = s:match("(%S)%s*$")
--						s = s:gsub("%s+$","")
						if j < #t.columns or last_ch:match("[!%+]") then
							s = s:match("(.+)%S%s*$")
							last_ch = s:sub(-1,-1)
							s = s:gsub("%s*$","")
						end				
						s = s:gsub("^%s+","")
						if cell.text == "" then						
							cell.text = s.."\r\n"
						else
							cell.text = cell.text..s.."\r\n"
						end	
						if j > col_index then
							cell.colspan = j - col_index + 1
						end	
						col_index = j + 1
					else	
						s = s:gsub("^%s+","")
						if cell.text == "" then						
							cell.text = s
						else
							cell.text = cell.text..s
						end	
					end	
				end
				j = j + 1
			until j > t.cols	
		end
		idx = idx + 1
	until idx > t.height
	start_row()
	t.rows = max_row() - 1
	for i = 1,t.rows do
		for j = 1, #t.columns do
			if cells[i] and cells[i][j] and cells[i][j].text then
				cells[i][j].text = cells[i][j].text:gsub("^[%s\r\n]+",""):gsub("[%s\r\n]+$",""):gsub("%s+\r\n","\r\n")
			end
		end
	end
	t.cells = cells	
	return t	
end

function render_grid(data)
	local text 
	local block = ""
	set_col_widths(data,"grid",CtrlWidth)
	for i = 1, data.rows do
		if not data.cells[i] then
			continue
		end
		local row_stru,row_data = "",""
		for j = 1, data.cols do
			local cell = data.cells[i][j]
			if not cell or cell.status and cell.status == "null" then 
				continue 
			end
			if not cell.status then
				cell.status = "normal"
			end	
			local cell_x = cell.x or data.columns[j].x
			if cell.colspan and not cell.rowspan then
				cell_x = data.columns[j+cell.colspan-1].x
				for k = 1, cell.colspan-1 do
					if data.cells[i][j+k] then
						data.cells[i][j+k].status = "null"
					end
				end
			end
			if cell.rowspan then
				cell.status = "clvmgf"
				for k = 1,cell.rowspan-1 do
					if not data.cells[i+k][j] then
						data.cells[i+k][j] = {}
					end
					local sub_cell = data.cells[i+k][j]
					sub_cell.status = "clvmrg"
					if cell.colspan then
						cell_x = data.columns[j+cell.colspan-1].x
						sub_cell.x = cell_x
						for m = 1, cell.colspan-1 do
							if data.cells[i+k][j+m] then
								data.cells[i+k][j+m].status = "null"
							end
						end
					end	
				end
			end
			local cell_stru = ""
			if cell.status == "clvmgf" then
				cell_stru = cell_stru..[[\clvmgf]]
			elseif cell.status == "clvmrg" then
				cell_stru = cell_stru..[[\clvmrg]]
			end
			for _, item in pairs({"l","t","r","b"}) do
				cell_stru = cell_stru..[[\clbrdr]]..item..[[\brdrs\brdrw10]]
			end	
			local cell_text = [[\pard\plain\intbl]]
			local text = ""
			if cell.status ~= "clvmrg" then
				if data.hdrH and i <= data.hdrH then
					cell_text = cell_text..[[\qc]]
					text = render_lines( {Style = "gridH", Lines = cell.text:split("\r\n")} )
					cell_stru = cell_stru..[[\clcbpat5\clvertalc]]
				else
					cell_text = cell_text..[[\]]..data.columns[j].align
					text = render_lines({Style = "grid", Lines = cell.text:split("\r\n")} )
				end	
			end	
			cell_text = cell_text..text..[[\cell]].."\r\n"
			cell_stru = cell_stru..[[\cellx]]..cell_x.."\r\n"
			row_stru = row_stru..cell_stru
			row_data = row_data..cell_text
		end	-- next column (j)
		block = block..[[\trowd\trautofit1\trgaph]]..GapH
		block = block..row_stru..row_data..[[\row]].."\r\n"
	end	-- next row (i)
	return block..[[\pard\par]]
end -- function render_grid

-------------------------------------------------------
--[[
	Анализ абзаца, определение вида таблицы (список/сетка), 
	формирование структуры с описанием таблицы
]]

function is_table(para)
	local lines = para.Lines
	local tp = ""
	local t = parse_table(lines)
	if t then
		return "table", t
	end	
	t = parse_grid(lines) 
	if t then
		return "grid", t
	end	
	return "" 
end

end 	-- tables

--[[
	Сборка RTF
]]

local function build_rtf(text)
	local rtf = {}
	setmetatable(rtf,{["__index"] = table})
	
--	таблица шрифтов	
	rtf:insert( [[{\fonttbl]] )
	rtf:insert( [[{\f0\fnil\default;}]] )
	for i=1,#RTF.Fonts do
		rtf:insert( ([[{\f%d\fnil %s;}]]):format(i,RTF.Fonts[i]) )
	end
	rtf:insert( "}\r\n" )
	
--	таблица цветов
	rtf:insert( [[{\colortbl;]] )
	for i=1,#RTF.Colors do
		local color = RTF.Colors[i]
		local r,g,b = color.R,color.G,color.B
		rtf:insert( ([[\c%d\red%d\green%d\blue%d;]]):format(i,r,g,b) )
	end
	rtf:insert( "}\r\n" )

	return "{"..RTF.Header..rtf:concat()..text .. "}"
end

local function MakeRTF(text, width)
	if width then
		CtrlWidth = width
	end
	local rtf = ToRTF(text)
	rtf = build_rtf(rtf)
	return rtf
end

get_dpi()

reMark = {
	makeRTF = MakeRTF,
	Styles = RTF.Styles,
}

return reMark