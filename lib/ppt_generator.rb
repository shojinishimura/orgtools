$: << File.dirname(__FILE__)
require "org_parser"
require "win32ole"

class PptGenerator
  def start
    @ppt = WIN32OLE.new("PowerPoint.Application")
    @ppt.Visible = true
    @doc = @ppt.Presentations.Add()
    @slide = nil
    
    @layout = {
      :blank => 12,
      :title => 1,
      :text => 2
    }
  end

  def finish
    
  end
  
  def do_headline(level, line, hash)
    case level
    when 0
      make_title_page(line)
    when 1
      make_new_page(line)
    end
  end

  def do_itemize(level, line)
    add_new_paragraph(line, :level => level + 1)
  end

  def do_contline(level, line)
    append_text(line)
  end
  
  def do_blankline
    
  end

  def do_normal(line)
    add_new_paragraph(line, :bullet_visible => false)
  end
  
  private
  def page_num
    @doc.Slides.Count
  end

  def new_slide(layout)
    @slide = @doc.Slides.Add(page_num + 1, @layout[layout])
  end

  def set_title(text)
    @slide.Shapes(1).TextFrame.TextRange.Text = text
  end

  # args
  #   :to default = @slide.Shapes(2).TextFrame.TextRange
  #   :level  indent level, default = 1
  #   :bullet_visible  bullet visibility, default = true
  def add_new_paragraph(text, args={ })
    to = args[:to] || @slide.Shapes(2).TextFrame.TextRange
    level = args[:level] || 1
    bullet_visible = args[:bullet_visible].nil? ? true : args[:bullet_visible]
    delimiter = to.Paragraphs.Count == 0 ? "" : "\r\n"
    
    append_text(text, :to => to, :delimiter => delimiter)

    paragraph = to.Paragraphs(to.Paragraphs.Count)
    paragraph.IndentLevel = level
    paragraph.ParagraphFormat.Bullet.Visible = bullet_visible
  end

  # args
  #   :to TextFrame object to append to, default = @slide.Shapes(2).TextFrame.TextRange
  #   :delimiter default = ""
  def append_text(text, args={ })
    to = args[:to] || @slide.Shapes(2).TextFrame.TextRange
    delimiter = args[:delimiter] || ""

    to.InsertAfter(delimiter)
    to.InsertAfter(text)
  end
  
  def make_title_page(line)
    ppLayoutTitle = 1
    new_slide(:title)
    set_title(line)
  end

  def make_new_page(line)
    ppLayoutText = 2
    new_slide(:text)
    set_title(line)
  end
end

if __FILE__ == $0
  parser = OrgParser.new(:visitor => PptGenerator.new)
  parser.parse($stdin)
end
