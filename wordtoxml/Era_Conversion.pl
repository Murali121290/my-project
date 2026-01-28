use strict;
use File::Copy;
use Cwd;
use Cwd 'abs_path';
use File::Basename;
use utf8;

my $InFile=$ARGV[0];
my $HtmFile=$ARGV[1];
print "\n$InFile";
#print "\n**********$HtmFile";

my ($HtmName, $HtmDir, $FileSuffix) = fileparse($HtmFile, "\.[a-z]+");


	my $cont=&HexToDec("$InFile", "HTML");
	
	#----------------------- Pre Clean ---------------------#
	$cont=~s#\n?<html[^>]*?>#<html>#gsi;
	$cont=~s#<p/>##gsi;
	$cont=~s#<br/>#\n#gsi;
	
	#-- Weblink
=head
	$cont=~s#([^"])(http\:[^\s<>]+|www[\.\,\s\;\:\-][^\s<>]+|[^\s<>]+(?:\.edu|\.com|\.org|\.gov|\.net|\.co|\.in))#$1<a href="$2">$2</a>#gsi;
	$cont=~s#<a href="(www)#<a href="http://$1#gsi;
=cut
	$cont=~s#<Hyperlink>(.*?)</Hyperlink>#&hyperlink($1)#gesi;
	sub hyperlink
	{
			my $hyper=shift;
			$hyper=~s#<a [^\>]*?>|</a>##gsi;
			$hyper="<a href=\"$hyper\">$hyper</a>";
			return $hyper;
	}
	
	$cont=~s#<a href="([^\"]*?)"#"<a href=\"".&ApplyHTTP($1)."\""#gesi;
	sub ApplyHTTP
	{
			my $href=shift;
			if ($href!~m#^http#si)
			{
					if ($href=~m#\@#si)
					{
						$href="mailto:$href";
					}
					else
					{
						$href="http://$href";
					}
			}
			return $href;
	}
	
	#-- Header Clean
	$cont=~s#\n?<html[^>]*?>#\n<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">\n<html xmlns="http://www.w3.org/1999/xhtml" >#gsi;
	$cont=~s#<title/>#<title></title>#gsi;
	$cont=~s#"../css/epub.css"#"epub.css"#gsi;
	
	#-- Page id
	$cont=~s#\&lt;pageid_(\w+)\/\&gt;#<a id="page\_$1"/>#gsi;
	$cont=~s#(?:<p>|<p [^\>]*?>)\s*(<a id="page\_\w+"/>)\s*</p>\s*(<p>|<p [^\>]*?>)#$2$1#gsi;
	
	#-- Formatting
	$cont=~s#<u>#<span class="Underline">#gsi;
	$cont=~s#</u>#</span>#gsi;
	
	
	#--------------------- Images -------------------#
	#-- Object and Image name
	my $HtmWitOutPunc=$HtmName;
	$HtmWitOutPunc=~s#\_##gsi;
	$cont=~s#<object/>Image(${HtmWitOutPunc}\d+)#<img src="img/$1\.jpg" class="" alt=""/>#gsi;
	
	#-- Object name
	# my $InlineimgCount='0001';
	# $cont=~s#<object/>#'<img src="img/'.$HtmName.'_inline_'.$InlineimgCount++.'.jpg" class="" alt=""/>'#gesi;

	#-- cimage style
	my $DispImgCount='0001';
	$cont=~s#<p class="cimage"/>#'<p class="cimage"><img src="img/'.$HtmName.'_'.$DispImgCount++.'.jpg" class="" alt=""/></p>'#gesi;
	
	#-- Remove empty para with class
	$cont=~s#<p class="[^\"]*?"/>##gsi;

	
	#------------------- Footnotes ------------------#
=head	
	#-- Footnote Ids
	my $FnCount='001';
	$cont=~s#(<p [^\>]*?class="footnote[^\"]*?\"[^\>]*?>)#"$1<a id=\"FN".$FnCount++."\"/>"#gesi;
	
	#-- Footnote Label move inside a tag
	$cont=~s#(<a id="FN\d+")/>\s*(\d+)#$1\>$2</a>#gsi;
	
	#-- Footnotes Linking - Page by Page
	$cont=~s#(<a id="page_\w+">)#</page>$1#gsi;
	$cont=~s#<a id="page_\w+"/>.*?(</page>|$)#&PageFootnote($&)#gesi;
	sub PageFootnote
	{
			my $fn=shift;
			my @FnArray;
			while ($fn=~m# id="(FN\d+)">(.*?)</a>#gsi)
			{
					push(@FnArray,"$1<>$2");
			}
			my $AllFNCont;
			my %AllFNitem;
			while($fn=~m#<a [^\>]*?id="(FN\d+)"[^\>]*>((?:(?!<a |</a>).)*?)</a>#gsi)
			{
					my ($aId,$aText)=("$1","$&");
					if (!$AllFNitem{$aId})
					{
						$AllFNitem{$aId}="$aText";
					}
					$AllFNCont.="$&\n";
			}
			$fn=~s#<a [^\>]*?id="(FN\d+)"[^\>]*>((?:(?!<a |</a>).)*?)</a>#<<$1>>#gsi;
			foreach (@FnArray)
			{
					my ($id,$text)=($1,$2) if ($_=~m#^(.*?)<>(.*?)$#si);
					$fn=~s#<sup>($text)</sup>#<a href="\#$id"><sup>$1</sup></a>#si;
			}
			$fn=~s#<<([^\>]*?)>>#$AllFNitem{$1}#gesi;
			return $fn;
	}
	$cont=~s#</page>##gsi;
	
	
	#-- Footnotes Movements
	my @footnotes;
	while ($cont=~m#<p [^\>]*?class="footnote[^\"]*?\"[^\>]*?>((?:(?!<p>|<p |</p>).)*?)</p>#gsi)
	{
		push(@footnotes,"$&");
	}
	$cont=~s#<p [^\>]*?class="footnote[^\"]*?\"[^\>]*?>((?:(?!<p>|<p |</p>).)*?)</p>\n?##gsi;

	my $footText=join("\n",@footnotes);
	if ($footText ne '')
	{
		$cont=~s#</body>#<p class="footline"></p>\n$footText\n</body>#gsi;
	}
=cut
	#-- List
	
	foreach ("bulletlist", "circlelist", "squarelist", "numberlist", "Alphauplist", "alphalolist", "romanuplist", "romanlolist")
	{
			my $style=$_;
			$cont=~s#((<p class=\s*"\s*($style)\s*"\s*>((?:(?!<p>|<p |</p>).)*?)</p>\s*)+)#&List_Conv($&,$style)#gesi;
	}
	
	sub List_Conv
	{
			my $list=shift;
			my $type=shift;
			$list=~s#<p [^\>]*?>#<li>#gsi;
			$list=~s#</p>#</li>#gsi;
			$list="<ul>\n".$list."\n</ul>" if ($type=~m#bulletlist#si);
			$list="<ul style=\"list-style-type:circle\">\n".$list."\n</ul>" if ($type=~m#circlelist#si);
			$list="<ul style=\"list-style-type:square\">\n".$list."\n</ul>" if ($type=~m#squarelist#si);
			$list="<ol type=\"1\">\n".$list."\n</ol>" if ($type=~m#numberlist#si);
			$list="<ol type=\"A\">\n".$list."\n</ol>" if ($type=~m#Alphauplist#si);
			$list="<ol type=\"a\">\n".$list."\n</ol>" if ($type=~m#alphalolist#si);
			$list="<ol type=\"I\">\n".$list."\n</ol>" if ($type=~m#romanuplist#si);
			$list="<ol type=\"i\">\n".$list."\n</ol>"if ($type=~m#romanlolist#si);
			return $list;
	}

	# &WriteFile("$HtmFile.tmp", "$cont", "HTML");
	# system "pause";
	#----------------------- Post Clean ----------------------#
	$cont=~s#</strong><strong>##gsi;
	$cont=~s#<p [^<>]*class=\"Book-Title\"[^<>]*>((?:(?!</p>|<p |<p>).)*)</p>#<h1 class=\"book-title\">$1<\/h1>#gsi;
	$cont=~s#<p [^<>]*class=\"book-subtitle\"[^<>]*>((?:(?!</p>|<p |<p>).)*)</p>#<h2 class=\"book-subtitle\">$1<\/h2>#gsi;
	$cont=~s#<p [^<>]*class=\"book-author\"[^<>]*>((?:(?!</p>|<p |<p>).)*)</p>#<h2 class=\"book-author\">$1<\/h2>#gsi;
	$cont=~s#<p [^<>]*class=\"(part-headers?|part-numbers?|part-names?)\"[^<>]*>((?:(?!</p>|<p |<p>).)*)</p>#<h2 class=\"$1\">$2<\/h2>#gsi;
	$cont=~s#<h2 class=\"(part-number|part-name|part-header)(s)\">#<h2 class=\"$1\">#isg;
	$cont=~s#<p [^<>]*class=\"Ch-header1\"[^<>]*>((?:(?!</p>|<p |<p>).)*)</p>#<h3 class=\"chapter-header-1\">$1<\/h3>#gsi;
	$cont=~s#<p [^<>]*class=\"Ch-header2\"[^<>]*>((?:(?!</p>|<p |<p>).)*)</p>#<h3 class=\"chapter-header-2\">$1<\/h3>#gsi;
	$cont=~s#<p [^<>]*class=\"Ch-header3\"[^<>]*>((?:(?!</p>|<p |<p>).)*)</p>#<h4 class=\"chapter-header-3\">$1<\/h4>#gsi;
	$cont=~s#<p [^<>]*class=\"Ch-header4\"[^<>]*>((?:(?!</p>|<p |<p>).)*)</p>#<h4 class=\"chapter-header-4\">$1<\/h4>#gsi;
	$cont=~s#<p [^<>]*class=\"Ch-header5\"[^<>]*>((?:(?!</p>|<p |<p>).)*)</p>#<h5 class=\"chapter-header-5\">$1<\/h5>#gsi;
	$cont=~s#<p [^<>]*class=\"Ch-header6\"[^<>]*>((?:(?!</p>|<p |<p>).)*)</p>#<h6 class=\"chapter-header-6\">$1<\/h6>#gsi;
	$cont=~s#<p [^<>]*class=\"Endmatter-header\"[^<>]*>((?:(?!</p>|<p |<p>).)*)</p>#<h3 class=\"endmatter-header\">$1<\/h3>#gsi;
	
	$cont=~s#class=\"(full-out|full-outDC)\"#class=\"full-out\"#gsi;
	$cont=~s#class=\"(right-align)\"#class=\"right\"#gsi;
	$cont=~s#class=\"(Hanging)\"#class=\"hanging-indent-full\"#gsi;
	
	$cont=~s#class=\"(Extract0SingleSpaced)\"#class=\"extract-FO-0-single-spaced\"#gsi;
	$cont=~s#class=\"(Extract0SglSpaced)\"#class=\"extract-FO-0-single-spaced\"#gsi;
	
	$cont=~s#class=\"(Extract-0-Digital)\"#class=\"extract-0-digital\"#gs;

	$cont=~s#class=\"(ExtractFO\-?0SingleSpaced)\"#class=\"extract-FO-0-single-spaced\"#gsi;
	$cont=~s#class=\"(ExtractFO\-?0SglSpaced)\"#class=\"extract-FO-0-single-spaced\"#gsi;
	$cont=~s#class=\"(ExtractFO\-?11st)\"#class=\"extract-FO-1-1st\"#gsi;
	$cont=~s#class=\"(ExtractFO\-?2mid)\"#class=\"extract-FO-2-mid\"#gsi;
	$cont=~s#class=\"(ExtractFO\-?3last)\"#class=\"extract-FO-3-last\"#gsi;
	
	$cont=~s#class=\"(ExtractIN-0mid|ExtractIN-0SingleSpaced)\"#class=\"extract-IN-0-single-spaced\"#gsi;
	$cont=~s#class=\"(ExtractIN-11st)\"#class=\"extract-IN-1-1st\"#gsi;
	$cont=~s#class=\"(ExtractIN-2mid)\"#class=\"extract-IN-2-mid\"#gsi;
	$cont=~s#class=\"(ExtractIN-3last)\"#class=\"extract-IN-3-last\"#gsi;

	$cont=~s#class=\"(ExtractRight1-1st)\"#class=\"extract-right-1-1st\"#gsi;
	$cont=~s#class=\"(ExtractRight2-mid)\"#class=\"extract-right-2-mid\"#gsi;
	$cont=~s#class=\"(ExtractRight3-last)\"#class=\"extract-right-3-last\"#gsi;
	$cont=~s#class=\"(ExtractRight)\"#class=\"right space-after-single\"#gsi;
	
	$cont=~s#class=\"(ExtractCentre0SingleSpaced)\"#class=\"extract-centre-0-single-spaced\"#gsi;
	$cont=~s#class=\"(ExtractCentre0SglSpaced)\"#class=\"extract-centre-0-single-spaced\"#gsi;
	$cont=~s#class=\"(ExtractCentre1SpaceBefore)\"#class=\"extract-centre-1-space-before\"#gsi;
	$cont=~s#class=\"(ExtractCentre2NoSpace)\"#class=\"extract-centre-2-no-space\"#gsi;
	$cont=~s#class=\"(ExtractCentre3SpaceAfter)\"#class=\"extract-centre-3-space-after\"#gsi;

	$cont=~s#class=\"(Verse0-FO1st)\"#class=\"verse-0-FO-1st\"#gsi;
	$cont=~s#class=\"(Verse1-FOmid)\"#class=\"verse-1-FO-mid\"#gsi;
	$cont=~s#class=\"(Verse2-INmid)\"#class=\"verse-2-IN-mid\"#gsi;
	$cont=~s#class=\"(Verse3-INlast)\"#class=\"verse-3-IN-last\"#gsi;
	$cont=~s#class=\"(Verse4-FOlast)\"#class=\"verse-4-FO-last\"#gsi;

	$cont=~s#class=\"(copyright-1-Space-Above)\"#class=\"copyright-1-space-above\"#gsi;
	$cont=~s#class=\"(copyright-2-NoSpace)\"#class=\"copyright-2-no-space\"#gsi;
	$cont=~s#class=\"(copyright-3-SpaceBelow)\"#class=\"copyright-3-space-below\"#gsi;
	
	$cont=~s#<p [^<>]*class=\"z-image\"[^<>]*>((?:(?!</p>|<p |<p>).)*)</p>#<div class="full-page-image">$1<\/div>#gsi;
	$cont=~s#<p ([^<>]*class=\"[^"]*\"[^<>]*)>\s*<(Sans-?serif|sans)>((?:(?!</p>|<p |<p>|<sans|<\/sans).)*)<\/\2>\s*</p>#<p class=\"extract-0-digital\">$3</p>#gsi;

	$cont=~s#class=\"z-Caption\"#class="caption"#gsi;
	
	$cont=~s#class=\"z-endnote-FO\"#class="endnote-full-out"#gsi;
	$cont=~s#class=\"z-endnote-IN\"#class="endnote-indent"#gsi;
	
	$cont=~s#class=\"z-FootnoteFO\"#class="footnote-full-out"#gsi;
	$cont=~s#class=\"z-FootnoteIN\"#class="footnote-indent"#gsi;
	
	$cont=~s#class=\"z-Notehang\"#class="note"#gsi;

	$cont=~s#<p [^<>]*class=\"z-list\-?numbered\"[^<>]*>((?:(?!</p>|<p |<p>).)*)</p>#<ol-numb><li>$1<\/li><\/ol-numb>#gsi;
	$cont=~s#<\/ol-numb>\s*<ol-numb>#\n#gsi;
	$cont=~s#<ol-numb>#<ol class=\"ordered-list\">#gsi;
	$cont=~s#<\/ol-numb>#<\/ol>#gsi;

	$cont=~s#<p [^<>]*class=\"z-list\-?bullet\"[^<>]*>((?:(?!</p>|<p |<p>).)*)</p>#<ul-bull><li>$1<\/li><\/ul-bull>#gsi;
	$cont=~s#<\/ul-bull>\s*<ul-bull>#\n#gsi;
	$cont=~s#<ul-bull>#<ul class=\"unordered-list\">#gsi;
	$cont=~s#<\/ul-bull>#<\/ul>#gsi;
	
	
	#Formatting
	$cont=~s#<(strong|Bold|underline|Strikethrough|Superscript|Subscript|SmallCaps|monospace|sans-?serif|Italic|u)/>##isg;
	$cont=~s#<(Bold|underline|Strikethrough|Superscript|Subscript|SmallCaps|monospace|sans-?serif)Italic>#<$1><italic>#isg;
	$cont=~s#<\/(Bold|underline|Strikethrough|Superscript|Subscript|SmallCaps|sans-?serif|SmallCaps|monospace|sans-?serif)Italic>#<\/italic><\/$1>#isg;

	$cont=~s#<em>#<italic>#isg;
	$cont=~s#<\/em>#<\/italic>#isg;

	$cont=~s#<sub>#<subscript>#isg;
	$cont=~s#<\/sub>#<\/subscript>#isg;

	$cont=~s#<sup>#<superscript>#isg;
	$cont=~s#<\/sup>#<\/superscript>#isg;
	
	$cont=~s#<monospace>#<mono>#isg;
	$cont=~s#<\/monospace>#<\/mono>#isg;

	$cont=~s#<sans-\?serif>#<sans>#isg;
	$cont=~s#<\/sans-\?serif>#<\/sans>#isg;
	
	$cont=~s#<(\/)?(SmallCaps)>#<$1small-caps>#isg;
	
	$cont=~s#<(Bold|Italic|underline|Strikethrough|Superscript|Subscript|sans|small-caps|mono)>#"<span class=\"".lc($1)."\">"#isge;
	$cont=~s#<\/(Bold|Italic|underline|Strikethrough|Superscript|Subscript|sans|small-caps|mono)>#<\/span>#isg;

	
	$cont=~s#\n+#\n#gsi;
 &WriteFile("$HtmFile", "$cont", "HTML");
	#	&WriteFile_DecToChar("$HtmFile", "$cont", "HTML");


#-------------------------------- Sub Functions -------------------------------#

sub ReadFile
{
	my ($infile, $type)=@_;
	open (IN,"<$infile") or die("Unable to open $type file $infile: $!");
	undef $/; my $cont=<IN>;
	close IN;
	return $cont;
}
sub ReadFileDec
{
	my ($infile, $type)=@_;
	open (IN,'<:utf8', "$infile") or die("Unable to open $type file $infile: $!");
	undef $/; my $cont=<IN>;
	close IN;
	$cont=~s#([^\x00-\x7F])#"\&\#".ord($1)."\;"#gesi;		#-- Char to Decimal
	return $cont;
}
sub ReadFileHex
{
	my ($infile, $type)=@_;
	open (IN,'<:utf8', "$infile") or die("Unable to open $type file $infile: $!");
	undef $/; my $cont=<IN>;
	close IN;
	$cont=~s#([^\x00-\x7F])#"\&\#x".sprintf("%04X",ord($1))."\;"#gesi;	#-- Char to Hexa Decimal
	return $cont;
}
sub DecToHex
{
	my ($infile, $type)=@_;
	open (IN,"<$infile") or die("Unable to open $type file $infile: $!");
	undef $/; my $cont=<IN>;
	close IN;
	$cont=~s#(\&\#(\d+);)#"\&\#x".sprintf('%04X', "$2")."\;"#gesi;
	return $cont;
}
sub HexToDec
{
	my ($infile, $type)=@_;
	open (IN,"<$infile") or die("Unable to open $type file $infile: $!");
	undef $/; my $cont=<IN>;
	close IN;
	$cont=~s#\&\#x(\w+);#"\&\#".hex($1)."\;"#gesi;
	return $cont;
}
sub WriteFile
{
	my $outfile=shift;
	my $cont=shift;
	my $type=shift;
	open (OUT,">$outfile") or die("Unable to write $type file $outfile: $!");
	print OUT $cont;
	close OUT;
}
sub WriteFile_DecToChar
{
	my $outfile=shift;
	my $cont=shift;
	my $type=shift;
	$cont=~s#\&\#(\d+);#chr($1)#gesi;
	open (OUT,'>:utf8', "$outfile") or die("Unable to write $type file $outfile: $!");
	print OUT $cont;
	close OUT;
}
sub WriteFile_HexToChar
{
	my $outfile=shift;
	my $cont=shift;
	my $type=shift;
	$cont=~s#\&\#x(\w+);#chr(hex($1))#gesi;
	open (OUT,'>:utf8', "$outfile") or die("Unable to write $type file $outfile: $!");
	print OUT $cont;
	close OUT;
}
