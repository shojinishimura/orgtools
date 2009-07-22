class OrgParser

  class DefaultVisitor
    def method_missing(meth, *args, &block)
      puts "#{meth}(#{args.map{|e| e.inspect}.join(",")})\n"
    end
  end
  
  def initialize(*args)
    conf = args[0] || Hash.new
    @visitor = conf[:visitor] || DefaultVisitor.new
  end
  
  def parse(src)
    @visitor.start
    src.each do |line|
      line.chomp!
      case line
      when /^(\*+) (.*)/ # headline
        level = $1.nil? ? 0 : $1.size - 1
        line = $2.gsub(/ (\[.*\])?\s*(:.*:)?$/, "")
        args = Hash.new
        args[:timestamp] = $1 unless $1.nil?
        args[:tag] = $2 unless $1.nil?
        @visitor.do_headline(level, line, args)
      when /^( *)[-+*] (.*) :: (.*)/ # description
        level = $1.nil? ? 0 : $1.size / 2
        name = $2
        line = $3
        @visitor.do_description(level, name, line)
      when /^( *)[-+*] (.*)/ # itemize
        level = $1.nil? ? 0 : $1.size / 2
        line = $2
        @visitor.do_itemize(level, line)
      when /^( *)([1-9][0-9]*)\.? (.*)/ # enumerate
        level = $1.nil? ? 0 : $1.size / 2
        num = $2.to_i
        line = $3
        @visitor.do_enumerate(level, num, line)
      when /^(  )+(.*)/
        level = $1.size / 2 - 1
        line = $2
        @visitor.do_contline(level, line)
      when /^\|-/
        @visitor.do_tablesep
      when /^\|(.*)\|\s*$/
        row = $1.split("|").map {|e| e.strip}
        @visitor.do_table(row)
      when /^\s*:END:/
        @visitor.do_drawerend
      when /^\s*:([_A-Z]+):/ # drawers
        name = $1
        @visitor.do_drawer(name)
      when /^\s*:(w+):(.*)/
        name = $1
        value = $2.strip
        @visitor.do_property(name, value)
      when /^\#\+(.*): (.*)/
        tag = $1
        args = $2
        @visitor.do_specialcomment(tag, args)
      when /^\#\s*(.*)/
        line = $1
        @visitor.do_comment(line)
      when ""
        @visitor.do_blankline
      else
        @visitor.do_normal(line)
      end
    end
    @visitor.finish
  end
end

if __FILE__ == $0
  parser = OrgParser.new
  parser.parse($stdin)
end
