use Archive::Zip qw/ :ERROR_CODES :CONSTANTS /;
use Archive::Zip;
use Cwd 'abs_path';
use Cwd;
use Encode qw(decode encode);
use File::Basename;
use File::Copy::Recursive qw(dircopy);
use File::Copy::Recursive qw(pathrmdir);
use File::Copy;
use File::Find;
use File::HomeDir;
use File::Spec;
use File::stat;
use HTTP::Tiny;
use List::MoreUtils qw( minmax );
use POSIX qw(strftime);
use strict;
use String::Substitution qw( sub_modify );
use Sys::Hostname;
use Try::Tiny;
#use Uniq;
use utf8;
use warnings;         # still get other warnings
no warnings 'uninitialized';   # but silence uninitialized warnings
# use Win32; # Commented out for Linux compatibility

$|=1;

my $ExePath=abs_path($0);
$ExePath =~ s|[\\/][^\\/]+$||; # Platform independent regex

opendir(my $dh, $ARGV[0]) or die $!;
my @docx = grep { /\.docx$/i && -f "$ARGV[0]/$_" } readdir($dh);
closedir $dh;

foreach my $file (@docx)
{
#========================== Declarations ==========================#
#exit;
		my $Doc_File=$ARGV[0] . "/" . $file;
		
		my $Client_Name="Amazon";
		
#=========================== Extract DOCX ==========================#
		my (@ID,@Label);
		my $docPath = dirname($Doc_File);

		my @suffixes=(".docx",".docx");
		my $FileName= basename($Doc_File, @suffixes);

		my $zipname = $Doc_File;

		my $File_Path=dirname(abs_path($0));
		# $File_Path=~s#\/#\\#gsi; # Removed backslash forcing

		my $Final_File=File::Spec->catfile($docPath, "html", "$FileName.xml");
		my $Comments=File::Spec->catfile($docPath, "html", "Comments.html");
		my $Footnotes=File::Spec->catfile($docPath, "html", "${FileName}_Footnotes.html");
		mkdir (File::Spec->catdir($docPath, "html")) if (!-d File::Spec->catdir($docPath, "html"));

		print "Converting to XML $FileName...\n";
#========================== Read ZIP File ==========================#

		my $zip = Archive::Zip->new($zipname);
		
		foreach my $member ($zip->members)
		{
			(my $extractName = $member->fileName) =~ s{.*/}{};

			my $extractName1 = $member->fileName;

			if($extractName eq "document.xml")
			{
				if($extractName1=~m{word/document.xml}gsi)
				{
					$member->extractToFileNamed("$docPath/$extractName");
					my $XML_File="$docPath/$extractName";
					my $Post_XML="$docPath/html/$FileName.posthtml";
					$Post_XML =~s/\.xml$/\.postxml/i;
					#					my $XMLCont=&ReadFile("$XML_File", "HTML");
					#					$XMLCont =~ s#<w:r ((?:(?!0Hidden|<w:r |<\/w:r>).)*)<w:rStyle w:val="0Hidden"\/>((?:(?!0Hidden|<w:r |<\/w:r>).)*)<\/w:r>##isg;
					#					$XMLCont =~ s#<w:fldChar w:fldCharType="begin"/>((?:(?!w:fldCharType="begin"|w:fldCharType="end").)*)<w:fldChar w:fldCharType="end"/>##isg;
					#					&WriteFile("$XML_File", "$XMLCont", "HTML");

					#					system("perl \"$File_Path\\Era_WmlCleanup.pl\" \"$XML_File\" \"$Doc_File\"");
					system("perl \"$File_Path/Era_WmlCleanup.pl\" \"$XML_File\" \"$Doc_File\"");
					system("java -jar \"$File_Path/saxon.jar\" \"$XML_File\" \"$File_Path/Era_Word2XML.xsl\" > \"$Post_XML\"");
					system("python3 \"$File_Path/utf8_converter.py\" \"$Post_XML\"");
					# system("$File_Path\\List.exe \"$Post_XML\" \"$Post_XML\"");
					#					print "\n$Post_XML => $Final_File\n";
					#					system("perl \"$File_Path\\Era_Conversion.pl\" \"$Post_XML\" \"$Final_File\" \"$Client_Name\"");
					system("perl \"$File_Path/Era_Conversion.pl\" \"$Post_XML\" \"$Final_File\" \"$Client_Name\"");
				unlink("$Post_XML");
				}
			}
=head
			if($extractName eq "footnotes.xml")
			{
				if($extractName1=~m{word\/footnotes.xml}gsi)
				{
					$member->extractToFileNamed("$docPath/$extractName");
					my $XML_File="$docPath/$extractName";
					my $Post_XML="$docPath/html/$FileName\_Footnotes.posthtml";
					$Post_XML =~s/\.xml$/\.postxml/i;

					system("perl \"$File_Path/Era_WmlCleanup.pl\" \"$XML_File\" \"$Doc_File\"");
					# system("\"$File_Path\\Era_WmlCleanup.exe\" \"$XML_File\" \"$Doc_File\"");
					system("java -jar \"$File_Path/saxon.jar\" \"$XML_File\" \"$File_Path/Era_Word2XML.xsl\" > \"$Post_XML\"");
					system("python3 \"$File_Path/utf8_converter.py\" \"$Post_XML\"");
					# system("$File_Path\\List.exe \"$Post_XML\" \"$Post_XML\"");

					system("perl \"$File_Path/Era_Conversion.pl\" \"$Post_XML\" \"$Footnotes\" \"$Client_Name\"");
					# system("\"$File_Path\\Era_Conversion.exe\" \"$Post_XML\" \"$Footnotes\" \"$Client_Name\"");
					unlink("$Post_XML");
				}
			}
			
			if($extractName eq "comments.xml")
			{
				if($extractName1=~m{word\/comments.xml}gsi)
				{
					$member->extractToFileNamed("$docPath/$extractName");
					my $XML_File="$docPath/$extractName";
					my $Post_XML="$docPath/html/Comments.posthtml";
					$Post_XML =~s/\.xml$/\.postxml/i;

					system("perl \"$File_Path/Era_WmlCleanup.pl\" \"$XML_File\" \"$Doc_File\"");
					# system("\"$File_Path\\Era_WmlCleanup.exe\" \"$XML_File\" \"$Doc_File\"");
					system("java -jar \"$File_Path/saxon.jar\" \"$XML_File\" \"$File_Path/Era_Word2XML.xsl\" > \"$Post_XML\"");
					system("python3 \"$File_Path/utf8_converter.py\" \"$Post_XML\"");
					# system("$File_Path\\List.exe \"$Post_XML\" \"$Post_XML\"");
					#					print "\n$Post_XML => $Comments\n";
					system("perl \"$File_Path/Era_Conversion.pl\" \"$Post_XML\" \"$Comments\" \"$Client_Name\"");
					# system("\"$File_Path\\Era_Conversion.exe\" \"$Post_XML\" \"$Comments\" \"$Client_Name\"");
					unlink("$Post_XML");
				}
			}
=cut
			if($extractName eq "custom.xml")
			{
					$member->extractToFileNamed("$docPath/$extractName");
					my $XML_File="$docPath/$extractName";
					my $Cust_XML="$docPath\\Custom1.xml";

					my ($Editor);
					my $Tmp=&ReadFile("$Final_File", "HTML");
					my $Tmp1=&ReadFile("$XML_File", "HTML");

					if($Tmp1=~m{<property ([^\>]+)\>(.*?)<\/property>}gsi)
					{
							my $Name=$1;
							my $Content=$2;
							if($Name=~m{name=\"editor\"}gsi)
							{
								$Content=~s{<vt:lpwstr>(.*?)<\/vt:lpwstr>}{}gsi;
								$Editor=$1;

								$Tmp=~s{<front>}{<\?CE $Editor\?>\n<front>}gsi;
							}
					}
					#					print "\n$Final_File";
					&WriteFile("$Final_File", "$Tmp", "HTML");
					unlink("$XML_File");
					unlink("$Cust_XML");
			}
		}
		copy("$File_Path\\epub.css", "$docPath\\html\\epub.css");
		rename("$docPath\\$FileName.zip",$Doc_File);
		unlink("$docPath/document.xml");
		unlink("$docPath/footnotes.xml");
		unlink("$docPath/$FileName\_Footnotes.html");
		unlink("$docPath/Comments.html");
		unlink("$docPath/document.posthtml");

#============================ XML File ============================#
my $booMeta=<<BKMETA;
<?xml version="1.0" encoding="UTF-8"?>
<!--<!DOCTYPE book PUBLIC "-//NLM//DTD BITS Book Interchange DTD v2.0 20130520//EN" "http://jats.nlm.nih.gov/extensions/bits/2.0/BITS-book2.dtd">-->
<book xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xi="http://www.w3.org/2001/XInclude" xmlns:mml="http://www.w3.org/1998/Math/MathML" xmlns:xlink="http://www.w3.org/1999/xlink" dtd-version="2.0" xml:lang="en">
<book-body>
<book-part id="ch&num;" book-part-type="chapter">
<book-part-meta>
<title-group>
<label>&ChLabel;</label>
<title>&ChTitle;</title>
</title-group>
<contrib-group>
&contrib;
</contrib-group>
<abstract>
<title>Abstract</title>
&abstract;
</abstract>
<kwd-group kwd-group-type="author">
<title>&KeyTermsHeading;</title>
&kwd;
</kwd-group>
</book-part-meta>
<body>
BKMETA
		
		my $Tmp=&ReadFile("$Final_File", "HTML");
		$Tmp=~s#(\&\#(\d+);)#"\&\#x".sprintf('%04X', "$2")."\;"#gesi;
		#-- cleanup
		#		print "\n$Final_File";
		$Tmp=~s#^.*?<body>##isg;
		$Tmp=~s#<\/body>\s*<\/html>#<\/body>\n</book-part>\n</book-body>\n</book>#isg;
		$Tmp=~s#(<\/?comment(?: [^<>]*)?>|<CommentReference[0-9]+\/>)##isg;
		#Formatting
		$Tmp=~s#&lt;\/?(bold|ital)&gt;##isg;
		$Tmp=~s#<span class="italic">((?:(?!</span>|<span ).)*?)</span>#<italic>$1<\/italic>#isg;
		$Tmp=~s#<\/?normaltextrun>##isg;
		while($Tmp=~s#<(italic|strong)>(\s*)</\1>#$2#isg){}
		while($Tmp=~s#<(italic|strong)>(\s*[\.\,]+\s*)</\1>#$2#isg){}
		$Tmp=~s#<\/(TableNumber|FigureNumber)>\.#\.<\/$1>#isg;
		$Tmp=~s#<\/strong>([^<>]*)<\/strong>#$1<\/strong>#isg;
		$Tmp=~s#<strong>([^<>]*)<strong>#<strong>$1#isg;

		while($Tmp=~s#<\/(strong|italic|citebib|bibchaptertitle|bibtitle|bibjournal|bibarticle|bibpublisher|biburl|bibvolume|bibissue|bibfpage|biblpage|bibsurname|bibfname|TableNumber|TableCitation|FigureCitation|FigureNumber)>(\s*)<\1>#$2#isg){}
		$Tmp=~s#<p(?: [^<>]*)?><strong>&lt;(\/)?(case study)&gt;</strong></p>#<$1casestudy>#isg;
		#		&WriteFile("$Final_File\.tmp", "$Tmp", "HTML");
		$Tmp=~s#&amp;#&\#x0026;#isg;
		
		my $num = "";
		if($Tmp=~s#<p class="ChapterNumber">\s*(?:<strong>)?(Chapter ([0-9]+))\s*(?:</strong>)?\s*</p>##is){
			my $lab = $1; $num = $2;
			$booMeta=~s#&ChLabel;#$lab#isg;
			$booMeta=~s#&num;#$num#isg;
		}
		if($Tmp=~s#<p class="ChapterTitle">\s*(?:<strong>)?((?:(?!</p>|<p ).)*?)\s*(?:<\/strong>)?\s*</p>##is){
			my $tit = $1;
			$booMeta=~s#&ChTitle;#$tit#isg;
		}
		if($Tmp=~s#<p class="PartNumber">\s*(?:<strong>)?((?:Section|Part) ([0-9A-Z\.\-]+))\s*(?:</strong>)?</p>\s*<p class="PartTitle">\s*(?:<strong>)?((?:(?!</p>|<p ).)*?)\s*(?:<\/strong>)?\s*</p>##is){
			my $lab = $1; my $id = $2; my $tit = $3;
			$booMeta=~s#<book-body>#<book-body>\n<book-part id="pt$id" book-part-type="part">\n<book-part-meta>\n<title-group>\n<label>$lab</label>\n<title>$tit</title>\n</title-group>\n</book-part-meta>\n<body>#isg;
			$Tmp=~s#<\/body>\s*</book-part>\s*</book-body>\s*</book>#</body>\n</book-part>\n</body>\n</book-part>\n</book-body>\n</book>#isg;
		}
		if($Tmp=~s#<p class="ChapterAuthor">((?:(?!<\/p>|<p ).)*?)<\/p>##is){
			my $chapauth = $1;
			$chapauth=~s#<[^>]*>##isg;
			$chapauth=~s#( and )#\n#isg;
			$chapauth=~s#(\s*\,\s*)#\n#isg;
			$chapauth=~s#^([^<>\n]+) ([^<> \n]+)$#<contrib contrib-type="author">\n<name>\n<surname>$2</surname> <given-names>$1</given-names>\n</name>\n</contrib>#img;
			$booMeta=~s#&contrib;#$chapauth#isg;
		}
		if($Tmp=~s#<p class="[^"]*">\s*(?:<strong>)*Abstract(?:<\/strong>)*\s*</p>((?:\s*<p class="[^"]*">((?:(?!<\/p>|<p ).)*?)<\/p>))##is){
			my $abs = $1;
			$abs=~s#<p class="[^"]*">#<p>#isg;
			$booMeta=~s#&abstract;#$abs#isg;
		}
		$Tmp=~s#<p class="[^"]*">\s*Key<strong>w</strong>ords\s*</p>#<p class="SP-Heading2">Keywords</p>#isg;
		if($Tmp=~s#<p class="[^"]*">\s*(?:<strong>)*(Keywords?)(?:<\/strong>)*\s*</p>\s*((?:\s*<p class="[^"]*">((?:(?!<\/p>|<p ).)*?)<\/p>))##is){
			my $tit = $1; my $kwd = $2;
			$booMeta=~s#&KeyTermsHeading;#$tit#isg;
			$kwd=~s#\s*<\/?p(?: [^<>]*)?>\s*##isg;
			if($kwd=~s#\s*\,\s*#<\/kwd>\n<kwd>#isg){
			}else{
				$kwd=~s#\s*\;\s*#<\/kwd>\n<kwd>#isg;
			}
			$booMeta=~s#&kwd;#<kwd>$kwd<\/kwd>#isg;
		}
	
		#casestudy
		$Tmp=~s#<casestudy>((?:(?!<\/casestudy>|<casestudy>).)*?)<\/casestudy>#caseStudy($&)#isge;
		
		$Tmp=~s#<p class="(Para-FL|ParaFirstLine-Ind)">#<p>#isg;
		$Tmp=~s#<p class="EOC-#<p class="#isg;
		
		#List Opener
		$Tmp=~s#<p class="BulletList([1-9])first0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<BL$1><p>$2</p><\/BL$1>#isg;
		$Tmp=~s#<p class="LearningObj-BulletList([1-9])-first0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<BL$1><p>$2</p><\/BL$1>#isg;
		#unList Opener
		$Tmp=~s#<p class="UnNumberList([1-9])first0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<UL$1><p>$2</p><\/UL$1>#isg;
		$Tmp=~s#<p class="UnNumberList([1-9])0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<UL$1><p>$2</p><\/UL$1>#isg;
		#BL List body
		$Tmp=~s#<p class="BulletList([1-9])0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<BL$1><p>$2</p><\/BL$1>#isg;
		$Tmp=~s#<p class="LearningObj-BulletList([1-9])0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<BL$1><p>$2</p><\/BL$1>#isg;
		#OL List body
		$Tmp=~s#<p class="NumberList([1-9])first0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<OL$1><p>$2</p><\/OL$1>#isg;
		$Tmp=~s#<p class="NumberList([1-9])0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<OL$1><p>$2</p><\/OL$1>#isg;
		$Tmp=~s#<p class="AlphaListfirst([1-9])0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<AL$1><p>$2</p><\/AL$1>#isg;
		$Tmp=~s#<p class="AlphaList([1-9])0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<AL$1><p>$2</p><\/AL$1>#isg;

		#		$Tmp=~s#<p class="(BulletList1first0|LearningObj-BulletList1-first)">((?:(?!<\/p>|<p ).)*?)<\/p>#listHead($&,$1)#isge;
		$Tmp=~s#<p class="(KeyTerm)">((?:(?!<\/p>|<p ).)*?)<\/p>#listHead($&,$1)#isge;
	
		#Keyterms
		$Tmp=~s#<\/kt1>\s*<kt1>#\n#isg;
		$Tmp=~s#<\/kt1>#\n<\/list>#isg;
		$Tmp=~s#<kt1>#<list list-type="bullet">#isg;
		my $i = 1;
		$Tmp=~s#&seq;#$i++#isge;
		#		$cont=~ s#<p class=\"(NL|UL|BL|OL|TOC-Chapter)([0-9]+)\">((?:(?!<p |<p>|<\/p>).)*)</p>#<$1$2>$3<\/$1$2>#isg;
		# Tags to be checked for nesting
		my @tags = qw(UL BL OL AL AU RL RU);

		# Primary nesting within same type
		foreach my $tag (@tags) {
		    $Tmp = fix_nesting($Tmp, $tag, $tag);
		}

		# Cross nesting (e.g., OL inside BL, NL inside UL, etc.)
		foreach my $outer (@tags) {
		    foreach my $inner (@tags) {
			next if $outer eq $inner; # skip same-type (already done)
			$Tmp = fix_nesting($Tmp, $outer, $inner);
		    }
		}
		
		$Tmp=~ s#<\/(UL|BL|OL|AL|AU|RL|RU)([0-9]+)>\s*<\1\2>#<\/list-item>\n<list-item>#isg;
		$Tmp=~ s#<\/(UL|BL|OL|AL|AU|RL|RU)([0-9]+)>#<\/list-item><\/$1$2>#isg;

		$Tmp=~ s#<(UL|BL|OL|AL|AU|RL|RU)([0-9]+)>#<ol class="$1">\n<list-item>#sg;
		$Tmp=~ s#<ol class="OL">#<list list-type="order">#isg;
		$Tmp=~ s#<ol class="AL">#<list list-type="lower-alpha">#isg;
		$Tmp=~ s#<ol class="AU">#<list list-type="upper-alpha">#isg;
		$Tmp=~ s#<ol class="RL">#<list list-type="lower-roman">#isg;
		$Tmp=~ s#<ol class="RU">#<list list-type="upper-roman">#isg;
		$Tmp=~ s#<ol class="BL">#<list list-type="bullet">#isg;
		$Tmp=~ s#<ol class="UL">#<list list-type="none">#isg;
		$Tmp=~ s#<\/(UL|BL|OL|AL|AU|RL|RU)([0-9]+)>#\n<\/list>\n#sg;
		$Tmp=~ s#<list-item><p>\s*<tab/>#<list-item><p>#isg;
		$Tmp=~ s#<list-item><p>\&\#x[0-9a-z]+;<tab/>#<list-item><p>#isg;
		$Tmp=~s#<p class="(References?Heading1|KeyTerms?Heading|LearnObjHeading)">((?:(?!<\/p>).)*?)<\/p>#<p class="Head1">$2<\/p>#isg;
		$Tmp=~s#<p class="(SpecialHeading)([0-9])">((?:(?!<\/p>).)*?)<\/p>#<p class="Head$2">$3<\/p>#isg;
		$Tmp=~s#<p class="Head(1|2|3|4|5|6)">((?:(?!<p |<\/p>).)*?)<\/p>#<sec$1 disp-level="level$1" id="ch${num}lev$1sec&seq1;">\n<title>$2<\/title>\n<\/sec$1>#gsi;
		$Tmp=~s#<title><strong>((?:(?!<strong>|<\/strong>).)*?)</strong></title>#<title>$1</title>#isg;
		$Tmp=~s#<title>\s*&lt;(KT|H[0-9]+)&gt;#<title>#isg;

		$Tmp=~s#^(.*?)$#SecLevel($&)#gsie;
		$Tmp=~s#<\/casec>#<\/sec>#gsi;
		$Tmp=~s#<casec #<sec #gsi;
		
		$i = 1;
		$Tmp=~s#&seq1;#$i++#isge;
		$Tmp=~s#<p class="Reference-Alphabetical">((?:(?!<p |<\/p>).)*?)<\/p>#"<ref-list>\n".ReferenceCode($&)."\n<\/ref-list>"#gsie;
		$i = 1;
		$Tmp=~s#&seq2;#$i++#isge;
		$Tmp=~s#<\/ref-list>\s*<ref-list>#\n#gsi;
		$i = 1;
		$Tmp=~s#&seq3;#$i++#isge;
		#figure
		$Tmp=~s#<p class="FigureLegend">\s*<strong>\s*(Figure [0-9\.\-]+) ((?:(?!<p |<\/p>).)*?)<\/strong>\s*<\/p>#<fig id="fig${num}_&seq3;" orientation="portrait" position="float"><label>$1</label>\n<caption><title>$2</title></caption>\n<graphic xmlns:xlink="http://www.w3.org/1999/xlink" orientation="landscape" xlink:href="media/.ai" mime-subtype="jpeg"/>\n</fig>#isg;
		$Tmp=~s#<p class="FigureLegend">\s*<FigureNumber>(?:<strong>)?\s*(Figure [0-9\.\-]+)(?:<strong>)?<\/FigureNumber>\s*<strong>((?:(?!<p |<\/p>).)*?)<\/strong>\s*<\/p>#<fig id="fig${num}_&seq3;" orientation="portrait" position="float"><label>$1</label>\n<caption><title>$2</title></caption>\n<graphic xmlns:xlink="http://www.w3.org/1999/xlink" orientation="landscape" xlink:href="media/.ai" mime-subtype="jpeg"/>\n</fig>#isg;
		$i = 1;
		$Tmp=~s#&seq3;#$i++#isge;
		#table
		$Tmp=~s#<table(?: [^<>]*)\/>#<table frame="box" rules="all" border="0" cellpadding="1" cellspacing="1">#gsi;
		$Tmp=~s#<table\/>#<\/table>#gsi;
		$Tmp=~s#<p class="TableCaption">\s*<strong>(Table [0-9\.\-]+) ((?:(?!<p |<\/p>).)*?)<\/strong>\s*</p>#<table-wrap id="tab${num}_&seq3;" position="float" orientation="portrait" content-type="table">\n<label>$1</label>\n<caption>\n<title>$2</title>\n</caption></table-wrap>#isg;
		$Tmp=~s#<p class="TableCaption">\s*<TableNumber>(?:<strong>)?(Table [0-9\.\-]+)(?:<\/strong>)?<\/TableNumber>\s*<strong>((?:(?!<p |<\/p>).)*?)<\/strong>\s*</p>#<table-wrap id="tab${num}_&seq3;" position="float" orientation="portrait" content-type="table">\n<label>$1</label>\n<caption>\n<title>$2</title>\n</caption></table-wrap>#isg;
		$Tmp=~s#<p class="TableSource">((?:(?!<p |<\/p>).)*?)</p>#<table-wrap><table-wrap-foot>\n<attrib>$1</attrib>\n</table-wrap-foot></table-wrap>#isg;
		$Tmp=~s#<table( [^<>]*)?>((?:(?!<table |<\/table>).)*?)<\/table>#tableClean($&)#gsie;
		$Tmp=~s#<\/table-wrap>\s*<table-wrap>#\n#gsi;
		$i = 1;
		$Tmp=~s#&seq3;#$i++#isge;
		#RefLink
		my $reflist = $Tmp;
		while($reflist=~ s#<ref ((?:(?!<ref |<\/ref>).)*)\n((?:(?!<ref |<\/ref>).)*)<\/ref>#<ref $1$2<\/ref>#isg){}
		while($reflist=~ s#<ref ((?:(?!<ref |<\/ref>).)*)&\#x(2014|2013|2011|2010|2012)\;((?:(?!<ref |<\/ref>).)*)<\/ref>#<ref $1\-$3<\/ref>#isg){}

		#		&WriteFile("$Final_File\.ref", "$reflist", "HTML");
		$Tmp=~ s#<citebib>((?:(?!</citebib>|<citebib>).)*)</citebib>#&refLinker($&,$reflist)#isge;
		while($Tmp=~ s#<p id="(term[0-9]+)">([^<>]*)</p>(.*?)<strong>(\2s?)</strong>#<p id="$1">$2</p>$3<bold><xref ref-type="keyterm" rid="$1">$4</xref></bold>#isg){}
		#Figure and Table Links
		my $Tmp1 = $Tmp;
	while($Tmp1=~s#<fig [^<>]*id=\"([^"]*)\"[^<>]*>(\s*)<label>((?:(?!<\/label>|<label>).)*)<\/label>##is){
		my $id = $1;
		my $Label = $3;
		$Label=~s#\s+$##isg;
		$Label=~s#^\s+##isg;
		push(@ID,"$id");
		push(@Label,"$Label");
	}

	while($Tmp1=~s#<table-wrap[^<>]*id=\"([^"]*)\"[^<>]*>(\s*)<label>((?:(?!<\/label>|<label>).)*)<\/label>##is){
		my $id = $1;
		my $Label = $3;
		$Label=~s#\s+$##isg;
		$Label=~s#^\s+##isg;
		push(@ID,"$id");
		push(@Label,"$Label");
	}
		while($Tmp=~s#<(FigureCitation|TableCitation)>([^<>]+)<\/\1>#<TLink label="$2">$2<\/TLink>#isg){
			
		}
LOOP:
		while($Tmp=~s#<TLink label="([^"<> ]*) ([^"<>]*)">((?:(?!<\/TLink>|<TLink ).)*)<\/TLink>([^<>]+)<TLink label="([^"<> ]+)">#<TLink label="$1 $2">$3<\/TLink>$4<TLink label="$1 $5">#isg){
			print "\n$&";
			goto LOOP;
		}
		#		print "\nL1";
		for(my $i = 0; $i<scalar(@Label);$i++){
		#			print "\n**$Label[$i]** => $ID[$i]";
			$Tmp=~s#<TLink label="\Q$Label[$i]\E">#<TLink href="$ID[$i]">#isg;
		}

		#		print "\nL2";
		for(my $i = 0; $i<scalar(@Label);$i++){
			$Tmp=~s#<TLink label="\Q$Label[$i]\E[a-z]">#<TLink href="$ID[$i]">#isg;
		}

		#		print "\nL3";
		$Tmp=~s#<xref ref-type="([^"]*)"><TLink href="([^"]*)">((?:(?!<\/TLink>|<Tlink ).)*?)<\/TLink>((?:(?!<\/xref>|<xref ).)*?)</xref>#<xref ref-type="$1" rid="$2">$3$4</xref>#isg;
		$Tmp=~s#<TLink href="([^"]*)">((?:(?!<\/TLink>|<Tlink ).)*?)<\/TLink>#<xref ref-type="" rid="$1">$2</xref>#isg;
		$Tmp=~s#<xref ref-type="[^"]*" rid="(fig[^"]*)">#<xref ref-type="fig" rid="$1">#isg;
		$Tmp=~s#<xref ref-type="[^"]*" rid="(tab[^"]*)">#<xref ref-type="table" rid="$1">#isg;
		my $finalCont = "$booMeta$Tmp";
		$finalCont=~s#\n\s*\n#\n#isg;

		# set floating elements into cited place
		my %float;
		
		# capture figures and tables
		while ($finalCont =~ m{(<(fig|table-wrap)\b[^>]*\bid="([^"]+)"[^>]*>.*?</\2>)}sgx) {
		    $float{$3} = $1;
		}

		# remove the block
		$finalCont =~ s{(<(fig|table-wrap)\b[^>]*\bid="([^"]+)"[^>]*>.*?</\2>)}{}sgx;
		
		foreach my $rid (keys %float) {
			$finalCont =~ s{(<p\b[^>]*>.*?<xref\b[^>]*\brid="$rid"[^>]*>.*?</p>)}{

			    my ($p) = ($1);

			    # If matching float exists
			    if (exists $float{$rid}) {

				my $block = $float{$rid};

				# Prevent double insertion
				delete $float{$rid};

				"$p\n$block";
			    }
			    else {
				# No citation leave paragraph unchanged
				$p;
			    }
			}xseg;
		}
		
		foreach my $rid (keys %float) {
			$finalCont =~ s{</ref-list>}{</ref-list>\n$float{$rid}}sg;
		}

		$finalCont=~s#(\n+)#\n#isg;
		&WriteFile("$Final_File", "$finalCont", "HTML");
#========================= Sub Functions =========================#

sub ReadFile
{
	my ($infile, $type)=@_;
	open (IN,"<$infile") or die("Unable to open $type file $infile: $!");
	undef $/; my $cont=<IN>;
	close IN;
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

sub listHead{
	my $tmp = shift;
	my $type = shift;
	if($type=~m#(BulletList1first|LearningObj-BulletList1-first)#is){
	#		$tmp=~s#<p class="(BulletList1first[0-9]+|LearningObj-BulletList1-first)">((?:(?!<\/p>|<p ).)*?)<\/p>#<list list-type="bullet">\n<list-item>\n<p>$2</p>\n</list-item><\/bl1>#isg;
	}else{
	#		$tmp=~s#<p class="(BulletList1|LearningObj-BulletList1)">((?:(?!<\/p>|<p ).)*?)<\/p>#<bl1>\n<list-item>\n<p>$2</p>\n</list-item><\/bl1>#isg;
		$tmp=~s#<p class="(KeyTerm)">((?:(?!<\/p>|<p ).)*?)<\/p>#<kt1>\n<list-item>\n<p id="term&seq;">$2</p>\n</list-item><\/kt1>#isg;
	}
	return $tmp;
}

sub SecLevel{
	my $lvlCont = shift;
	$lvlCont=~s#</sec\d+>##gsi;
	$lvlCont=~s#<\/body>#<sec1><empty><\/body>#is;
	$lvlCont=~s#(<sec\d+)#<enter>$1#gsi;
	my @body;
	my ($lvl,$prevlvl);
	$lvl = $prevlvl = 0;
	@body = split ('<enter>', $lvlCont);
	foreach my $line (@body)
	{
		if($line =~ /<sec(\d+)[^\>]*?>/)
		{
			$lvl=ord($1);
			my $closetag = '';
			for(my $i=$prevlvl;$i>=$lvl;$i--)
			{
				my $l = chr($i);
				$closetag = $closetag . "\n</sec$l>";
			}
			$line = $closetag . "\n" . $line;
			$prevlvl = $lvl;
		}
	}

	my $lines =join("\n",@body);
	$lvlCont=$lines;
	$lvlCont=~s#\n{1,}#\n#gsi;
	$lvlCont=~s#\n(<\/sec\d+>)#$1#gsi;
	$lvlCont=~s#<sec(\d+)><empty>##gsi;
	$lvlCont=~s#<(\/)?sec(\d+)#<$1sec#gsi;
	return $lvlCont;
}

sub fix_nesting {
    my ($text, $outer_tag, $inner_tag) = @_;
    for my $i (reverse 1..6) {
        my $j = $i + 1;
        next if $j > 6;
        while ($text =~ s#</$outer_tag$i>\n*(<${inner_tag}$j>((?:(?!<${inner_tag}$j>).)*?)</${inner_tag}$j>)#\n$1\n</$outer_tag$i>#gsi) {}
    }
    return $text;
}
sub ReferenceCode{
    my $text = shift;
    $text=~s#<\/?Untag>##isg;
    
    #Authors
    $text=~s#<(\/)?biborganization>#<$1collab>#isg;
    $text=~s#<(bibsurname|bibfname)>((?:(?!<\/\1>|<\1>).)*?)<\/\1>#<pg><strname>$&</strname></pg>#isg;
    $text=~s#</pg>([^<>]*)<pg>#$1#isg;
    $text=~s#<\/bibsurname></strname>([^<>a-z]*)<strname><bibfname>#<\/bibsurname>$1<bibfname>#isg;
    $text=~s#<(\/)?bibsurname>#<$1surname>#isg;
    $text=~s#<(\/)?bibfname>#<$1given-names>#isg;
    $text=~s#<(\/)?strname>#<$1string-name>#isg;
    $text=~s#<pg>#<person-group person-group-type="author">#isg;
    $text=~s#<\/pg>#<\/person-group>#isg;
    
    #year
    $text=~s#<bibyear>((?:(?!<\/bibyear>|<bibyear>).)*?)<\/bibyear>#yearFix($1)#isge;
    
    #titles
    $text=~s#<(\/)?bibchaptertitle>#<$1chapter-title>#isg;
    $text=~s#<(\/)?bibtitle>#<$1source>#isg;
    
    $text=~s#<(\/)?bibarticle>#<$1article-title>#isg;
    $text=~s#<(\/)?bibjournal>#<$1source>#isg;

    $text=~s#<(\/)?bibpublisher>#<$1publisher-name>#isg;
    $text=~s#<(\/)?bib(volume|issue|fpage|lpage)>#<$1$2>#isg;
    $text=~s#<volume><italic>(.*?)</italic></volume>#<volume>$1</volume>#isg;
    #url
    $text=~s#<biburl>((?:(?!<\/biburl>|<biburl>).)*?)<\/biburl>#<ext-link ext-link-type="uri" xlink:href="$1">$1</ext-link>#isg;
    $text=~s#<bibdoi>((?:(?!<\/bibdoi>|<bibdoi>).)*?)<\/bibdoi>#<ext-link ext-link-type="doi" xlink:href="$1">$1</ext-link>#isg;
    
    if($text=~m#(<\/ext-link>)#is){
	    $text=~s#<p class="Reference-Alphabetical">#<ref id="bid_${num}_&seq2;"><mixed-citation publication-type="web">#isg;
    }elsif($text=~m#(<\/issue>|<\/article-title>)#is){
	    $text=~s#<p class="Reference-Alphabetical">#<ref id="bid_${num}_&seq2;"><mixed-citation publication-type="article">#isg;
    }elsif($text=~m#(<\/chapter-title>|<\/publisher-name>|<\/publisher-loc>)#is){
	    $text=~s#<p class="Reference-Alphabetical">#<ref id="bid_${num}_&seq2;"><mixed-citation publication-type="book">#isg;
    }else{
	    $text=~s#<p class="Reference-Alphabetical">#<ref id="bid_${num}_&seq2;"><mixed-citation publication-type="other">#isg;
    }
    $text=~s#<\/p>#<\/mixed-citation><\/ref>#isg;
    #    $text=~s#><#>\n<#isg;
    return $text;
}
sub yearFix{
    my $text = shift;
    $text=~s#([a-z][a-z]+)#<month>$1<\/month>#isg;
    $text=~s#([0-9]+)#<day>$1<\/day>#isg;
    $text=~s#<day>([0-9][0-9][0-9][0-9])<\/day>([a-z]?)#<year>$1$2<\/year>#isg;
    return $text;
}
sub caseStudy{
    my $text = shift;
    $text=~s#<p class="(CaseStudyTitle)">#<p class="Head0">#isg;
    $text=~s#<p class="CaseStudy-#<p class="#isg;
    $text=~s#<p class="H(?:ead)?(2|3|4|5|6)">((?:(?!<p |<\/p>).)*?)<\/p>#"<sec$1 disp-level=\"level".($1-1)."\" id=\"ch${num}lev$1sec&seq1;\">\n<title>$2<\/title>\n<\/sec$1>"#gsie;
    $text=~s#<p class="H(?:ead)?(0|1|2|3|4|5|6)">((?:(?!<p |<\/p>).)*?)<\/p>#<sec$1 disp-level="level$1" id="ch${num}lev$1sec&seq1;">\n<title>$2<\/title>\n<\/sec$1>#gsi;
    $text=~s#</casestudy>#</body></casestudy>#isg;
    $text=~s#^(.*?)$#SecLevel("$&")#gsie;
    $text=~s#<\/body>##gsi;
    if($text=~s#<casestudy>\s*<sec [^<>]*disp-level="level0"[^<>]*>\s*<title>(?:<strong>)*((?:(?!<title>|<\/title>).)*?)(?:<\/strong>)*<\/title>#<boxed-text id="cs${num}_&seq3;" content-type="case study" position="float">\n<caption><title>$1</title></caption>#isg){
    #	    print "$&";
	    $text=~s#<title>\s*&lt;(KT|H[0-9]+)&gt;#<title>#isg;
	    $text=~s#<caption><title>\s*(Case Study) ([0-9\.\-]+:?)\s*#<label>$1 $2</label><caption><title>#isg;
	    $text=~s#<\/sec>\s*<\/casestudy>#<\/boxed-text>#isg;
	    $text=~s#<\/casestudy>#<\/boxed-text>#isg;
    }
    $text=~s#<\/sec>#<\/casec>#gsi;
    $text=~s#<sec #<casec #gsi;
    return $text;
}
sub tableClean{
    my $text = shift;
    $text=~s#<tgroup [^<>]*\/>##gsi;
    $text=~s#<colspec [^<>]*colnum="([^"]+)"[^<>]*\/>#<colgroup>\n<col content-type="$1"/>\n</colgroup>#gsi;
    $text=~s#<\/colgroup>\s*<colgroup>#\n#gsi;
    $text=~s#<entry([^<>]*)( colname="[^"]*")([^<>]*)>#<td$1$3>#gsi;
    $text=~s#<\/entry>#<\/td>#gsi;
    $text=~s#<row([^<>]*)>#<tr>#gsi;
    $text=~s#<\/row>#<\/tr>#gsi;
    $text=~s#<tbody>((?:(?!<tbody |<\/tbody>).)*?TableColumnHead(?:(?!<tbody |<\/tbody>).)*?)<\/tbody>#tablethead($&)#gsie;
    $text=~s##<p>#isg;
    $text=~s#<\/p>\s*<p class="TableBody[0-9]*">#<br\/>#isg;
    $text=~s#(<\/p>|<p(?: [^<>]*)?>)##isg;
    $text=~s#\s+<\/t(h|d)>#<\/t$1>#isg;
    $text=~s#</tbody>\s*<tbody>##isg;
    return "<table-wrap>$text</table-wrap>";
}
sub tablethead{
    my $text = shift;
    $text=~s#<td( [^<>]*)?>#<th$1>#gsi;
    $text=~s#<\/td>#<\/th>#gsi;
    $text=~s#<tbody>((?:(?!<tbody |<\/tbody>).)*?TableColumnHead(?:(?!<tbody |<\/tbody>).)*?)<\/tbody>#<thead>$1<\/thead>#gsi;
    $text=~s#<p class="TableColumnHead[0-9]*">#<p>#isg;
    return $text;
}
sub refLinker{
	my $tmp = shift;
	my $ref = shift;
	my $refCall = $tmp;
	#	print "\n$refCall";
	$tmp=~ s#<[^<>]*>##isg;
	$tmp=~ s#(et al|et al\.|\&amp\;|\&\#x0026\;)#&#isg;
	$tmp=~ s# and # & #isg;
	$tmp=~ s#(\,|\)|\(|\.)#&#isg;
	$tmp=~ s#\##\\\##isg;
	$tmp=~ s# #&#isg;
	my $tt = $tmp;
	my $yr = $1 if($tt=~ m#([0-9][0-9][0-9][0-9][a-z]?)$#img);
	while($tmp=~ s#\&\s*\&#&#isg){}
	$tmp=~ s#\&#.*?#img;
	my $rep = $ref=~ s#<ref ([^<>]*)>\s*<mixed-citation[^<>]*>(?:\s*<[^<>]*>\s*)*\s*$tmp((?:(?!<mixed-citation |<\/mixed-citation>).)*)<\/mixed-citation>#$&#img;
	#	my $rep2 = $ref=~ s#<ref ([^<>]*)>((?:(?!<ref <\/ref>).)*)$tt.*?$yr((?:(?!<ref <\/ref>).)*)<\/ref>#$&#img;
	if($rep == 1){
		if($ref=~ m#<ref [^<>]*id=\"([^"]*)\"[^<>]*>\s*<mixed-citation[^<>]*>(?:\s*<[^<>]*>\s*)*\s*$tmp((?:(?!<mixed-citation |<\/mixed-citation>).)*)<\/mixed-citation>#im){
			my $id = $1;
			$refCall=~ s#<citebib>#<xref ref-type="bibr" rid="$id">#isg;
			$refCall=~ s#<\/citebib>#<\/xref>#isg;
		}
	}else{
	#		print "\n$rep => $tmp";
			$refCall=~ s#<citebib>#<nocitebib>#isg;
			$refCall=~ s#<\/citebib>#<\/nocitebib>#isg;
	}
	return $refCall;
}

#</citebib>
#<volume><italic>24</italic></volume>
#ext-link-type="doi"
}

print "\nProcess Completed Successfully!\n";
# Win32::MsgBox("Process Completed Successfully!",0,"S4C");
