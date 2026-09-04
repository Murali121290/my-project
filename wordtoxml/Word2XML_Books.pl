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
use XML::LibXML;

my $LabelMap = {
    fig   => { re => qr/Fig(?:ur(?:e|es))?s?/i,  reftype => 'fig',        prefix => 'fig' },
    table => { re => qr/Tab(?:le|les)?s?\.?/i,    reftype => 'table',      prefix => 'tab' },
    box   => { re => qr/Box(?:es)?/i,             reftype => 'boxed-text', prefix => 'box' },
    video => { re => qr/Video(?:s)?/i,            reftype => 'video',      prefix => 'vid' },
    exhibit => { re => qr/Exhibit(?:s)?/i,            reftype => 'boxed-text',      prefix => 'exhibit' },
    casestudy => { re => qr/Case Study(?:s)?/i,            reftype => 'casestudy',      prefix => 'cs' },
};

# internal chapter.figure separator: "." or non-breaking hyphen ONLY
# (plain "-" is reserved for the range connector - see note below)
my $NUM     = qr/\d+(?:[.\x{2011}]\d+)?/;
my $SUFFIX  = qr/[A-Za-z]?/;
my $ITEM    = qr/${NUM}${SUFFIX}/;
my $CONNECT = qr/(?:\s*[\x{2013}-]\s*|\s+(?:and|through|to)\s+|,\s*(?:and\s+)?|\s+&\s+)/i;


$|=1;
my $ExePath=abs_path($0);
$ExePath=~s#[\\\/]([^\/\\]+)$##isg;

my $PYTHON_BIN = "python3";
for my $candidate ("python3", "python") {
	`$candidate --version` ;
	if ($? == 0) { $PYTHON_BIN = $candidate; last; }
}

opendir(my $dh, $ARGV[0]) or die $!;
my @docx = grep { /\.docx$/i && -f "$ARGV[0]/$_" } readdir($dh);
closedir $dh;

foreach my $file (@docx)
{
#========================== Declarations ==========================#
#exit;


		my $Doc_File=$ARGV[0] . "/" . $file;

                my $commentdocx = $Doc_File;
                
                $commentdocx=~ s{\.docx}{\_comments\.docx};

		my $Client_Name="Amazon";
		
#=========================== Extract DOCX ==========================#
		my (@ID,@Label);
		my $docPath = dirname($commentdocx);

		my @suffixes=(".docx",".docx");
		my $FileName= basename($commentdocx, @suffixes);
                $FileName =~ s{_comments}{};
		my $zipname = $commentdocx;

		my $File_Path=dirname(abs_path($0));

                system($PYTHON_BIN, "$File_Path/process_comments.py", '-i', $Doc_File, '-o', $commentdocx);
                
		# $File_Path=~s#\/#\\#gsi;

		my $Final_File="$docPath/html/$FileName.xml";
		my $Final_ClassFile="$docPath/html/$FileName" . "_class.xml";
		my $Comments="$docPath/html/Comments.html";
		my $Footnotes="$docPath/html/$FileName\_Footnotes.html";
		mkdir ("$docPath/html") if (!-d "$docPath/html");

		print "Converting to XML $FileName...\n";
#========================== Read ZIP File ==========================#

		my $zip = Archive::Zip->new($zipname);
		
		foreach my $member ($zip->members)
		{
			(my $extractName = $member->fileName) =~ s{.*/}{};

			my $extractName1 = $member->fileName;
                        
			if($extractName eq "styles.xml")
			{
                                $member->extractToFileNamed("$docPath/$extractName");
                        }

			if($extractName eq "document.xml")
			{
				if($extractName1=~m{word/document.xml}gsi)
				{
					$member->extractToFileNamed("$docPath/$extractName");
					my $XML_File="$docPath/$extractName";
					my $Post_XML="$docPath/html/$FileName.posthtml";
					$Post_XML =~s/\.xml$/\.postxml/i;

					#					system("perl \"$File_Path/Era_WmlCleanup.pl\" \"$XML_File\" \"$Doc_File\"");
					system("perl \"$File_Path/Era_WmlCleanup.pl\" \"$XML_File\" \"$commentdocx\"");
					system("java -jar \"$File_Path/saxon.jar\" \"$XML_File\" \"$File_Path/Era_Word2XML.xsl\" > \"$Post_XML\"");
					system($PYTHON_BIN, "$File_Path/utf8_converter.py", $Post_XML);
					# system("$File_Path/List.exe \"$Post_XML\" \"$Post_XML\"");
					#					print "\n$Post_XML => $Final_File\n";
					#					system("perl \"$File_Path/Era_Conversion.pl\" \"$Post_XML\" \"$Final_File\" \"$Client_Name\"");
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

					# system("perl \"$File_Path\\Era_WmlCleanup.pl\" \"$XML_File\" \"$Doc_File\"");
					system("\"$File_Path\\Era_WmlCleanup.exe\" \"$XML_File\" \"$Doc_File\"");
					system("java -jar \"$File_Path\\saxon.jar\" \"$XML_File\" \"$File_Path\\Era_Word2XML.xsl\" > \"$Post_XML\"");
					system("\"$File_Path\\UTF8.exe\" \"$Post_XML\"");
					# system("$File_Path\\List.exe \"$Post_XML\" \"$Post_XML\"");

					#					 system("\"perl $File_Path\\Era_Conversion.pl\" \"$Post_XML\" \"$Footnotes\" \"$Client_Name\"");
					 system("\"$File_Path\\Era_Conversion.exe\" \"$Post_XML\" \"$Footnotes\" \"$Client_Name\"");
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

					#					system("perl \"$File_Path\\Era_WmlCleanup.pl\" \"$XML_File\" \"$Doc_File\"");
					system("\"$File_Path\\Era_WmlCleanup.exe\" \"$XML_File\" \"$Doc_File\"");
					system("java -jar \"$File_Path\\saxon.jar\" \"$XML_File\" \"$File_Path\\Era_Word2XML.xsl\" > \"$Post_XML\"");
					system("\"$File_Path\\UTF8.exe\" \"$Post_XML\"");
					# system("$File_Path\\List.exe \"$Post_XML\" \"$Post_XML\"");
					#					print "\n$Post_XML => $Comments\n";
					#					 system("perl \"$File_Path\\Era_Conversion.pl\" \"$Post_XML\" \"$Comments\" \"$Client_Name\"");
					 system("\"$File_Path\\Era_Conversion.exe\" \"$Post_XML\" \"$Comments\" \"$Client_Name\"");
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
					$Tmp=~ s{(<comment[^>]*>|</comment[^>]*>)}{}gi;
					$Tmp=~ s{(<comment[^>]*>|</comment[^>]*>)}{}gi;
					$Tmp=~ s#<p([^>]*)>((?:(?!<tab\/>).)*?)<tab\/>#<p$1>#ig;
					$Tmp=~ s#\&lt;LO\&gt;##ig;
					$Tmp=~ s#\&lt;SH([0-9]*)\&gt;##ig;
					$Tmp=~ s#\&lt;H([0-9]*)\&gt;##ig;
					&WriteFile("$Final_File", "$Tmp", "HTML");
                                        #&WriteFile("$Final_ClassFile", "$Tmp", "HTML");
					unlink("$XML_File");
					unlink("$Cust_XML");
			}
		}
		copy("$File_Path/epub.css", "$docPath/html/epub.css");
		rename("$docPath/$FileName.zip",$commentdocx);
		rename($commentdocx, $Doc_File);
		unlink("$docPath/document.xml");
		unlink("$docPath/footnotes.xml");
		unlink("$docPath/$FileName\_Footnotes.html");
		unlink("$docPath/Comments.html");
		unlink("$docPath/document.posthtml");
		
my $DTDPath = $ExePath;
$DTDPath =~ s{\\}{\/}g;

#============================ XML File ============================#
my $booMeta=<<BKMETA;
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE book PUBLIC "-//NLM//DTD BITS Book Interchange DTD v2.0 20130520//EN" "$DTDPath/BITS-Book-1.0-DTD/BITS-book1.dtd">
<book xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xi="http://www.w3.org/2001/XInclude" xmlns:mml="http://www.w3.org/1998/Math/MathML" xmlns:xlink="http://www.w3.org/1999/xlink" dtd-version="1.0" xml:lang="en">
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

#basic clean up

#$booMeta =~ s#<contrib-group>(\s*)\&contrib;(\s*)</contrib-group>#<contrib-group></contrib-group>#;
		
		my $Tmp=&ReadFile("$Final_File", "HTML");
		$Tmp=~s#(\&\#(\d+);)#"\&\#x".sprintf('%04X', "$2")."\;"#gesi;
		$Tmp=~s{([^\x00-\x7F])}{sprintf('&#x%04X;', ord($1))}ge;
		#-- cleanup
		#		print "\n$Final_File";
		$Tmp=~s#^.*?<body>##isg;
		$Tmp=~s#<\/body>\s*<\/html>#<\/body>\n</book-part>\n</book-body>\n</book>#isg;
		$Tmp=~s#(<\/?comment(?: [^<>]*)?>|<CommentReference[0-9]+\/>)##isg;
                $Tmp =~ s{
                    ((?:<p\s+class="Epigraph">.*?</p>\s*)+)
                    ((?:<p\s+class="EpigraphSource">.*?</p>\s*)+)?
                }{
                    my $paragraphs    = $1;
                    my $sources_block = $2 // ""; # Default to empty string if no source present

                    # Clean up standard epigraph paragraphs
                    $paragraphs =~ s/<p\s+class="Epigraph">/<p>/g;

                    # Convert each EpigraphSource paragraph to an <attrib> element if sources exist
                    if ($sources_block) {
                        $sources_block =~ s/<p\s+class="EpigraphSource">(.*?)<\/p>/<attrib>$1<\/attrib>/gs;
                    }

                    # Reconstruct output block
                    "<disp-quote content-type=\"epigraph\">\n"
                    . $paragraphs
                    . $sources_block
                    . "</disp-quote>"
                }gesx;

		$Tmp=~s# class="Box-01-BulletList1# class="BulletList1#igs;
		$Tmp=~s# class="TB-BulletList1# class="BulletList1#igs;
		$Tmp=~s# class="TB-BulletList2# class="BulletList2#igs;
		$Tmp=~s# class="Box-01-BulletList2# class="BulletList2#igs;
		$Tmp=~s# class="Exhibit-BulletList2# class="BulletList2#igs;
		$Tmp=~s# class="FE-01-BulletList1# class="BulletList1#igs;
		$Tmp=~s# class="Box-01-NumberList1# class="NumberList1#igs;
		$Tmp=~s# class="TB-NumberList1# class="NumberList1#igs;
		$Tmp=~s# class="FE-01-NumberList1# class="NumberList1#igs;
		$Tmp=~s# class="Box-01-UL-FL1# class="UnNumberList1#igs;
		$Tmp=~s# class="UL-FL1# class="UnNumberList1#igs;
		$Tmp=~s# class="TB-UL-FL1# class="UnNumberList1#igs;
                $Tmp=~s# class="Exhibit-UN-TableSource"# class="TableSource"#igs;
		$Tmp=~s#<p><Hyperlink/></p>##igs;
		$Tmp=~s#<a href="http://;">;</a>#;#igs;
                $Tmp =~ s{((?:<p\s+class="BulletList1[_]?(?:first|last)?">(.*?)</p>\s*(?:<p\s+class="ListItem-Para-FL1">(.*?)</p>\s*)?)+)}{
                    my $block = $1;
                    my $list_items = "";

                    while ($block =~ m{<p\s+class="BulletList1[_]?(?:first|last)?">(.*?)</p>\s*(?:<p\s+class="ListItem-Para-FL1">(.*?)</p>)?}gs) {
                        my $title = $1;
                        my $desc  = $2;

                        $list_items .= "<list-item><p>$title</p>\n";
                        if (defined $desc) {
                            $list_items .= "<p>$desc</p>\n";
                        }
                        $list_items .= "</list-item>\n";
                    }

                    "<list list-type=\"bullet\">\n" . $list_items . "</list>"
                }gesx;
                    
                $Tmp =~ s{((?:<p\s+class="BulletList1[_]?(?:first|last)?">(.*?)</p>\s*(?:<p\s+class="ListItemPara-FL1">(.*?)</p>\s*)?)+)}{
                    my $block = $1;
                    my $list_items = "";

                    while ($block =~ m{<p\s+class="BulletList1[_]?(?:first|last)?">(.*?)</p>\s*(?:<p\s+class="ListItemPara-FL1">(.*?)</p>)?}gs) {
                        my $title = $1;
                        my $desc  = $2;

                        $list_items .= "<list-item><p>$title</p>\n";
                        if (defined $desc) {
                            $list_items .= "<p>$desc</p>\n";
                        }
                        $list_items .= "</list-item>\n";
                    }

                    "<list list-type=\"bullet\">\n" . $list_items . "</list>"
                }gesx;

                $Tmp=~s{<a (.*)">((?:(?!<\/a>).)*?)<\/a>}{

                                my $tag = $&;
                                $tag =~ s{<a }{\@\@uri }g;
                                $tag =~ s{</a>}{\@\@/uri\@\@}g;
                                $tag =~ s{<([^>]*)>}{}g;
                                $tag =~ s{ href=}{ xlink:href=}g;
                                $tag =~ s{\@\@uri}{<ext-link}g;
                                $tag =~ s{\@\@/uri\@\@}{</ext-link>}g;
                                if ($tag =~ m{doi.org/})
                                {
                                        $tag =~ s{<ext-link }{<ext-link ext-link-type="doi" }g;
                                }
                                else
                                {
                                        $tag =~ s{<ext-link }{<ext-link ext-link-type="uri" }g;
                                }
                                qq($tag);

                }gie;
                $Tmp=~s{<p[^>]*></p>\n}{}g;
                $Tmp=~s{<span class="([^"]+)">(\s*)</span>}{$2}g;
                $Tmp=~s{&lt;KT&gt;}{<xref><bold>}g;
                $Tmp=~s{&lt;/KT&gt;}{</bold></xref>}g;
                $Tmp=~s{ ( +)}{ }g;

		#Formatting
		$Tmp=~s#&lt;\/?(bold|ital)&gt;##isg;
		$Tmp=~s#<span class="italic">((?:(?!</span>|<span ).)*?)</span>#<italic>$1<\/italic>#isg;
		$Tmp=~s#<span class="superscript">((?:(?!</span>|<span ).)*?)</span>#<sup>$1<\/sup>#isg;
		$Tmp=~s#<span class="subscript">((?:(?!</span>|<span ).)*?)</span>#<sub>$1<\/sub>#isg;
		$Tmp=~s#<span class="Underline">((?:(?!</span>|<span ).)*?)</span>#<underline>$1<\/underline>#isg;
		$Tmp=~s#<span class="small-caps">((?:(?!</span>|<span ).)*?)</span>#<sc>$1<\/sc>#isg;
		$Tmp=~s#<\/?normaltextrun>##isg;
		while($Tmp=~s#<(italic|strong)>(\s*)</\1>#$2#isg){}
		while($Tmp=~s#<(italic|strong)>(\s*[\.\,]+\s*)</\1>#$2#isg){}
		$Tmp=~s#<\/(TableNumber|FigureNumber|BoxNumber)>\.#\.<\/$1>#isg;
		$Tmp=~s#<\/strong>([^<>]*)<\/strong>#$1<\/strong>#isg;
		$Tmp=~s#<strong>([^<>]*)<strong>#<strong>$1#isg;

		while($Tmp=~s#<\/(strong|italic|citebib|bibchaptertitle|bibtitle|bibjournal|bibarticle|bibpublisher|biburl|bibvolume|bibissue|bibfpage|biblpage|bibsurname|bibfname|TableNumber|TableCitation|FigureCitation|FigureNumber|Box-01-BoxTitle)>(\s*)<\1>#$2#isg){}

                $Tmp=~s#<p[^>]*>(<strong>)?&lt;(\/)?(case study)&gt;(<\/strong>)?<\/p>#<$2casestudy>#isg;
                $Tmp=~s#<p[^>]*>&lt;(<strong>)?(\/)?(case study)(<\/strong>)?&gt;<\/p>#<$2casestudy>#isg;
                $Tmp=~s#<p[^>]*>(<strong>)?&lt;(\/)?(metadata)&gt;(<\/strong>)?<\/p>##isg;
                $Tmp=~s#<p[^>]*>(<strong>)?&lt;(\/)?(original to the au)&gt;(<\/strong>)?<\/p>##isg;
		#		&WriteFile("$Final_File\.tmp", "$Tmp", "HTML");
                $Tmp =~ s{&(?!amp;|lt;|gt;|quot;|apos;|#[0-9]+;|#x[0-9a-fA-F]+;|num;|ChLabel;|ChTitle;|contrib;|abstract;|kwd;|KeyTermsHeading;)}{&#x0026;}isg;
                $Tmp =~ s{&amp;}{\&\#x0026;}isg;
                
                my $num = "";
		if($Tmp=~s#<p class="ChapterNumber">\s*(?:<strong>)?(Chapter)?\s*([0-9]+)\s*(?:</strong>)?\s*</p>##is){
			my $lab = $1 . " " . $2; $num = $2;
                        $num =~ s/^\s+|\s+$//g;
                        $lab =~ s/^\s+|\s+$//g;
                        $booMeta=~s#&ChLabel;#$lab#isg;
			$booMeta=~s#&num;#$num#isg;
		}
		if($Tmp=~s#<p class="ChapterTitle">\s*(?:<strong>)?((?:(?!</p>|<p ).)*?)\s*(?:<\/strong>)?\s*</p>##is){
			my $tit = $1;
			$tit =~ s#(<strong>|</strong>)##g;
			$tit =~ s#(<CommentReference1>|</CommentReference1>)##g;
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
			$chapauth=~s#^([^<>\n]+) ([^<> \n]+)(\s*)$#<contrib contrib-type="author">\n<name>\n<surname>$2</surname> <given-names>$1</given-names>\n</name>\n</contrib>#img;
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

                #--------------- Learning Objectives block ---------------#
                $Tmp=~s{<p class=\"LearnObjHeading\">((?:(?!<\/p>).)*?)<\/p>\s*(?:<p class=\"LearnObj-Para-FL\">((?:(?!<\/p>).)*?)<\/p>\s*)?((?:<p class=\"LearnObj-NumberList1(?:first|last)?0?\">(?:(?!<\/p>).)*?<\/p>\s*)+)}{&LearnObjectives($1,$2,$3,$num)}gesi;
                $Tmp=~s{<p class=\"LearnObjHeading\">((?:(?!<\/p>).)*?)<\/p>\s*(?:<p class=\"LearnObj-Para-FL\">((?:(?!<\/p>).)*?)<\/p>\s*)?((?:<p class=\"LearnObj-BulletList1(?:first|last)?0?\">(?:(?!<\/p>).)*?<\/p>\s*)+)}{&LearnObjectives($1,$2,$3,$num)}gesi;
                $Tmp=~s{<p class=\"LearnObjHeading\">((?:(?!<\/p>).)*?)<\/p>\s*(?:<p class=\"LearnObj-Para-FL\">((?:(?!<\/p>).)*?)<\/p>\s*)?((?:<p class=\"LearnObj-UL-FL1(?:first|last)?0?\">(?:(?!<\/p>).)*?<\/p>\s*)+)}{&LearnObjectives($1,$2,$3,$num)}gesi;
                
		#casestudy
		$Tmp=~s#<casestudy>((?:(?!<\/casestudy>|<casestudy>).)*?)<\/casestudy>#caseStudy($&)#isge;
                $Tmp=~s{((?:<p class=\"CaseStudy-BulletList1(?:first|last)?0?\">(?:(?!<\/p>).)*?<\/p>\s*)+)}{&CaseStudyBulletList($1,$num)}gesi;
                $Tmp=~s{((?:<p class=\"CaseStudy-NumberList1(?:first|last)?0?\">(?:(?!<\/p>).)*?<\/p>\s*)+)}{&CaseStudyBulletList($1,$num)}gesi;
                $Tmp=~s{((?:<p class=\"CaseStudy-UL-FL1(?:first|last)?0?\">(?:(?!<\/p>).)*?<\/p>\s*)+)}{&CaseStudyUnNumberList($1,$num)}gesi;
                $Tmp=~s{<p\s+class="[^"]*Uc-RomanList1[^"]*">(.*?)</p>}{<RL1>$1</RL1>}gsi;
                $Tmp=~s{<p\s+class="[^"]*BulletList1[^"]*">(.*?)</p>}{<BL1>$1</BL1>}gsi;
                $Tmp=~s{((?:<(?:RL|BL)1>.*?</(?:RL|BL)1>\s*)+)}{&NestLists($1)}gesx;
                $Tmp=~s{((?:<p class=\"BulletList2[_]?(?:first|last)?0?\">(?:(?!<\/p>).)*?<\/p>\s*)+)}{&BulletList($1,$num)}gesi;
                $Tmp=~s{((?:<p class=\"BulletList1[_]?(?:first|last)?0?\">(?:(?!<\/p>).)*?<\/p>\s*)+)}{&CaseStudyBulletList($1,$num)}gesi;
                $Tmp=~s{((?:<p class=\"NumberList1[_]?(?:first|last)?0?\">(?:(?!<\/p>).)*?<\/p>\s*)+)}{&CaseStudyNumberList($1,$num)}gesi;
                $Tmp=~s{((?:<p class=\"UL-FL1[_]?(?:first|last)?0?\">(?:(?!<\/p>).)*?<\/p>\s*)+)}{&CaseStudyNumberList($1,$num)}gesi;
                $Tmp=~s{((?:<p class=\"Exhibit-UL-FL1(?:first|last)?0?\">(?:(?!<\/p>).)*?<\/p>\s*)+)}{&CaseStudyUnNumberList($1,$num)}gesi;

		$Tmp=~s#<p class="(Para-FL|Key-Para-FL|ParaFirstLine-Ind)">#<p>#isg;
		$Tmp=~s#<p class="EOC-#<p class="#isg;
		
		#List Opener
		$Tmp=~s#<p class="BulletList([1-9])first0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<BL$1><p>$2</p><\/BL$1>#isg;
		$Tmp=~s#<p class="LearnObj-BulletList([1-9])-first0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<BL$1><p>$2</p><\/BL$1>#isg;
		#unList Opener
		$Tmp=~s#<p class="UnNumberList([1-9])first0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<UL$1><p>$2</p><\/UL$1>#isg;
		$Tmp=~s#<p class="UnNumberList([1-9])0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<UL$1><p>$2</p><\/UL$1>#isg;
		$Tmp=~s#<p class="UnNumberList([1-9])last0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<UL$1><p>$2</p><\/UL$1>#isg;
		#BL List body
		$Tmp=~s#<p class="BulletList([1-9])0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<BL$1><p>$2</p><\/BL$1>#isg;
		$Tmp=~s#<p class="LearnObj-BulletList([1-9])0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<BL$1><p>$2</p><\/BL$1>#isg;
		#OL List body
		$Tmp=~s#<p class="NumberList([1-9])first0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<OL$1><p>$2</p><\/OL$1>#isg;
		$Tmp=~s#<p class="NumberList([1-9])0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<OL$1><p>$2</p><\/OL$1>#isg;
		$Tmp=~s#<p class="AlphaListfirst([1-9])0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<AL$1><p>$2</p><\/AL$1>#isg;
		$Tmp=~s#<p class="AlphaList([1-9])0?">((?:(?!<\/p>|<p ).)*?)<\/p>#<AL$1><p>$2</p><\/AL$1>#isg;
		$Tmp=~s#<p class="Uc-AlphaList([1-9])first">((?:(?!<\/p>|<p ).)*?)<\/p>#<AL$1><p>$2</p><\/AL$1>#isg;
		$Tmp=~s#<p class="Uc-AlphaList([1-9])">((?:(?!<\/p>|<p ).)*?)<\/p>#<AL$1><p>$2</p><\/AL$1>#isg;
		$Tmp=~s#<p class="Uc-AlphatList([1-9])first">((?:(?!<\/p>|<p ).)*?)<\/p>#<AL$1><p>$2</p><\/AL$1>#isg;
		$Tmp=~s#<p class="Uc-AlphatList([1-9])last">((?:(?!<\/p>|<p ).)*?)<\/p>#<AL$1><p>$2</p><\/AL$1>#isg;
		$Tmp=~s#<p class="Uc-AlphatList([1-9])">((?:(?!<\/p>|<p ).)*?)<\/p>#<AL$1><p>$2</p><\/AL$1>#isg;
		$Tmp=~s#<p class="Lc-AlphaList([1-9])">((?:(?!<\/p>|<p ).)*?)<\/p>#<AL$1><p>$2</p><\/AL$1>#isg;
		$Tmp=~s#<p class="Lc-RomanList([1-9])">((?:(?!<\/p>|<p ).)*?)<\/p>#<AL$1><p>$2</p><\/AL$1>#isg;
                $Tmp=~s{<\/list-item>(\s*)<\/list>(\s*)<list list-type="bullet2">((?:(?!<\/list>).)*?)<\/list>(\s*)<list list-type="([^"]+)">(\s*)<list-item>}{<list list-type="bullet">$3<\/list>$4<\/list-item>\n<list-item>}gs;
                $Tmp=~s{<\/list-item>(\s*)<\/list>(\s*)<list list-type="bullet2">((?:(?!<\/list>).)*?)<\/list>}{<list list-type="bullet">$3<\/list>\n<\/list-item>\n</list>}gs;
                $Tmp=~s{ list-type="bullet2"}{ list-type="bullet"}g;
                ## Disp-quote
		$Tmp=~s#<p class="eXtractTxt">((?:(?!<\/p>|<p ).)*?)<\/p>#<disp-quote><p>$1</p><\/disp-quote>#isg;

		#		$Tmp=~s#<p class="(BulletList1first0|LearnObj-BulletList1-first)">((?:(?!<\/p>|<p ).)*?)<\/p>#listHead($&,$1)#isge;
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
		$Tmp=~s#<\/boxsec>#<\/sec>#gsi;
		$Tmp=~s#<boxsec #<sec #gsi;
                #glossary
                $Tmp =~ s{
                    <p\s+class="GlossaryHeading">(.*?)</p>\s*
                    ((?:(?!<p\s+class="GlossaryTermDefinition).)*?)
                    ((?:<p\s+class="GlossaryTermDefinition(?:UL-FL1[^"]*)?">.*?</p>\s*)+)
                }{
                    my $title       = $1;
                    my $intro       = $2;
                    my $terms_block = $3;
                    my $def_list    = "";

                    # Match the main term definition, followed by any optional UL-FL1 list paragraphs
                    while ($terms_block =~ m{
                        <p\s+class="GlossaryTermDefinition">(.*?)</p>\s*
                        ((?:<p\s+class="GlossaryTermDefinitionUL-FL1[^"]*">.*?</p>\s*)*)
                    }gsx) {
                        my $content    = $1;
                        my $list_items = $2;

                        # Split term and definition on the en-space entity &#x2002;
                        my ($term, $def) = split(/\s*&#x2002;\s*/, $content, 2);

                        # Process the extra list paragraphs through your subroutine if they exist
                        my $formatted_list = "";
                        if ($list_items && $list_items =~ /\S/) {
                            $formatted_list = GlossaryUnNumberList($list_items);
                        }

                        $def_list .= "<def-item>\n";
                        $def_list .= "  <term>$term</term>\n";
                        $def_list .= "  <def>\n";
                        $def_list .= "    <p>$def</p>\n";
                        $def_list .= $formatted_list if $formatted_list; # Insert the list here
                        $def_list .= "  </def>\n";
                        $def_list .= "</def-item>\n";
                    }

                    # Reconstruct output block
                    "<glossary>\n"
                    . "  <title>$title</title>\n"
                    . $intro
                    . "  <def-list>\n"
                    . $def_list
                    . "  </def-list>\n"
                    . "</glossary>";

                }gesx;


		$i = 1;
                $Tmp=~s#<bibed\-#<bibed#g;
                $Tmp=~s#</bibed\-#</bibed#g;
		$Tmp=~s#&seq1;#$i++#isge;
		$Tmp=~s#<p class="ReferenceAlphabetical">#<p class="Reference-Alphabetical">#isg;
		$Tmp=~s#<p class="Reference-Alphabetical">((?:(?!<p |<\/p>).)*?)<\/p>#"<ref-list>\n".ReferenceCode($&)."\n<\/ref-list>"#gsie;
		$Tmp=~s#<p class="Reference-Numbered">((?:(?!<p |<\/p>).)*?)<\/p>#"<ref-list>\n".ReferenceCode($&)."\n<\/ref-list>"#gsie;
                $Tmp=~s#<ext-link ext-link-type="[^"]*" xlink:href=" ">(\s*)</ext-link>#$1#g;
		$i = 1;
		$Tmp=~s#&seq2;#$i++#isge;
		$Tmp=~s#<\/ref-list>\s*<ref-list>#\n#gsi;
		$i = 1;
		$Tmp=~s#&seq3;#$i++#isge;

                # Convert unlinked URLs into <ext-link> tags while excluding trailing quotes/punctuation
                $Tmp =~ s{
                    \b(https?://[^\s<>"{}|\\^`]+[^\s<>"{}|\\^`.,;:!])([\.,;:!]?)(?=\s|<|$)
                }{
                    my $url   = $1;
                    my $punct = $2;

                    # Determine link type (doi vs standard uri)
                    my $type = ($url =~ m{doi\.org/}) ? "doi" : "uri";

                    "<ext-link ext-link-type=\"$type\" xlink:href=\"$url\">$url</ext-link>$punct";
                }gexi;


                my $figurenumber = $FileName;
                my @bookid = split('_', $figurenumber);
                $figurenumber = $bookid[0];
                $figurenumber=~ s{([0-9]+)}{}g;
                if (length($num) == 1)
                {
                        $figurenumber = $figurenumber . "_F0" . $num;
                }
                else
                {
                        $figurenumber = $figurenumber . "_F" . $num;
                }
                
                #figure
		#$Tmp=~s#<p class="FigureLegend">\s*<strong>\s*(Figure)(\s*)([0-9\.\-]+)([0-9]*)((?:(?!<p |<\/p>).)*?)<\/strong>\s*<\/p>(\s*)(<p class="FigureSource">)?((?:(?!<p |<\/p>).)*?)(</p>)?(<p class="FigureNote">)?((?:(?!<p |<\/p>).)*?)(</p>)?#<fig id="fig${num}_&seq3;" orientation="portrait" position="float"><label>$1$2$3$4</label>\n<caption><title>$5</title></caption>\n<graphic xmlns:xlink="http://www.w3.org/1999/xlink" orientation="landscape" xlink:href="$figurenumber\-$3.eps" mime-subtype="jpeg"/>\n$6$7$8\n$9$10$11\n</fig>#isg;
		$Tmp=~s{<p class="FigureLegend">\s*(<strong>)?\s*(Figure)(\s*)([0-9]+)([\.\-0-9]*)(<\/strong>)?((?:(?!<p |<\/p>).)*?)(<\/strong>)?<\/p>}{
                my $tag = $&;
                my $fignum = $5;
                $fignum =~ s{\.}{}g;
                if (length($fignum) == 1)
                {
                        $fignum = "0" . $fignum;
                }
                $tag=~ s{<p class="FigureLegend">\s*(<strong>)?\s*(Figure)(\s*)([0-9]+)([\.\-0-9]*)(<\/strong>)?((?:(?!<p |<\/p>).)*?)(<\/strong>)?<\/p>}{<fig id="fig${num}_&seq3;" orientation="portrait" position="float"><label>$2$3$4$5</label>\n<caption><title>$7</title></caption>\n<graphic xmlns:xlink="http://www.w3.org/1999/xlink" orientation="landscape" xlink:href="$figurenumber\-$fignum.eps" mime-subtype="jpeg"/>\n</fig>}g;
                qq($tag);

}isge;
		$Tmp=~s{<p class="FigureLegend">\s*<FigureNumber>(?:<strong>)?\s*(Figure)(\s*)([0-9]+)([\.\-0-9]*)(?:<strong>)?<\/FigureNumber>\s*<strong>((?:(?!<p |<\/p>).)*?)<\/strong>\s*<\/p>}{
                        my $tag = $&;
                        my $fignum = $5;
                        $fignum =~ s{\.}{}g;
                        if (length($fignum) == 1)
                        {
                                $fignum = "0" . $fignum;
                        }
                        $tag=~ s{<p class="FigureLegend">\s*<FigureNumber>(?:<strong>)?\s*(Figure)(\s*)([0-9]+)([\.\-0-9]*)(?:<strong>)?<\/FigureNumber>\s*<strong>((?:(?!<p |<\/p>).)*?)<\/strong>\s*<\/p>}{<fig id="fig${num}_&seq3;" orientation="portrait" position="float"><label>$1$2$3$4</label>\n<caption><title>$5</title></caption>\n<graphic xmlns:xlink="http://www.w3.org/1999/xlink" orientation="landscape" xlink:href="$figurenumber\-$fignum.eps" mime-subtype="jpeg"/>\n</fig>}g;
                        qq($tag);
}isge;
                while (($Tmp =~ m{<\/fig>(\s*)<p class="FigureSource">}si) || ($Tmp =~ m{<\/fig>(\s*)<p class="FigureNote">}si))
                {
                        $Tmp=~ s{</fig>(\s*)<p class="FigureSource">((?:(?!</p>).)*?)</p>}{<attrib>$2</attrib></fig>$1}s;
                        $Tmp=~ s{</fig>(\s*)<p class="FigureNote">((?:(?!</p>).)*?)</p>}{<p>$2</p></fig>$1}s;
                }
		$i = 1;
		$Tmp=~s#&seq3;#$i++#isge;
		#table
		$Tmp=~s#<table(?: [^<>]*)\/>#<table frame="box" rules="all" border="0" cellpadding="1" cellspacing="1">#gsi;
		$Tmp=~s#<table\/>#<\/table>#gsi;
		$Tmp=~s#<p class="TableCaption">\s*(<strong>)?(Table)(\s*)([0-9]+)([\.\-0-9]*)\s*(<\/strong>)?((?:(?!<p |<\/p>).)*?)(<\/strong>)?</p>#<table-wrap id="tab${num}_&seq3;" position="float" orientation="portrait" content-type="table">\n<label>$2$3$4$5</label>\n<caption>\n<title>$7</title>\n</caption></table-wrap>#isg;
		$Tmp=~s#<p class="TableCaption">\s*<TableNumber>(?:<strong>)?(Table)(\s*)([0-9]+)([\.\-0-9]*)(?:<\/strong>)?<\/TableNumber>\s*<strong>((?:(?!<p |<\/p>).)*?)<\/strong>\s*</p>#<table-wrap id="tab${num}_&seq3;" position="float" orientation="portrait" content-type="table">\n<label>$1$2$3$4</label>\n<caption>\n<title>$5</title>\n</caption></table-wrap>#isg;
		$Tmp=~s#<p class="TableSource">((?:(?!<p |<\/p>).)*?)</p>#<table-wrap><table-wrap-foot>\n<attrib>$1</attrib>\n</table-wrap-foot></table-wrap>#isg;
                $Tmp=~s#</table>(\s*)<p class="TableNote">((?:(?!<p |<\/p>).)*?)</p>#<fn><p>$2</p></fn></table>#gs;
		$Tmp=~s#<table( [^<>]*)?>((?:(?!<table |<\/table>).)*?)<\/table>#tableClean($&)#gsie;
		$Tmp=~s#<\/table-wrap>\s*<table-wrap>#\n#gsi;
		$i = 1;
		$Tmp=~s#&seq3;#$i++#isge;

                #RefLink
		my $reflist = $Tmp;
		while($reflist=~ s#<ref ((?:(?!<ref |<\/ref>).)*)\n((?:(?!<ref |<\/ref>).)*)<\/ref>#<ref $1$2<\/ref>#isg){}
		while($reflist=~ s#<ref ((?:(?!<ref |<\/ref>).)*)&\#x(2014|2013|2011|2010|2012)\;((?:(?!<ref |<\/ref>).)*)<\/ref>#<ref $1\-$3<\/ref>#isg){}

		#		&WriteFile("$Final_File\.ref", "$reflist", "HTML");
                $Tmp=~ s{<mixed-citation[^<>]*>((?:(?!<\/mixed-citation>).)*)<\/mixed-citation>}{
                        my $tag = $&;
                        $tag=~ s{(<citebib>|<\/citebib>)}{}gs;
                        qq($tag);
                }gesi;
                $Tmp=~ s{<span class="([^"]+)">((?:(?!<\/span>).)*)</span>}{$2}g;
                $Tmp=~ s#<p class="ListHeading1">#<p content-type="flushleft">#g;
                $Tmp=~ s#<citebib>((?:(?!</citebib>).)*)</citebib>#&refLinker($&,$reflist)#isge;
                $Tmp=~ s#</nocitebib>\(<nocitebib>#\(#g;
                $Tmp=~ s#</nocitebib>\)#\)</nocitebib>#g;
                $Tmp=~ s#\(<nocitebib>#<nocitebib>\(#g;
                $Tmp=~ s#<nocitebib>(\s*)(\(|\))(\s*)</nocitebib>#$1$2$3#g;
                
                $Tmp = resolve_nocitebib($Tmp);
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
=cut        
		while($Tmp=~s#<(FigureCitation|TableCitation|BoxCitation)>([^<>]+)<\/\1>#<TLink label="$2">$2<\/TLink>#isg){
			
		}
LOOP:
		while($Tmp=~s#<TLink label="([^"<> ]*) ([^"<>]*)">((?:(?!<\/TLink>|<TLink ).)*)<\/TLink>([^<>]+)<TLink label="([^"<> ]+)">#<TLink label="$1 $2">$3<\/TLink>$4<TLink label="$1 $5">#isg){
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
=cut
		#		print "\nL3";
		#$Tmp=~s#<xref ref-type="([^"]*)"><TLink href="([^"]*)">((?:(?!<\/TLink>|<Tlink ).)*?)<\/TLink>((?:(?!<\/xref>|<xref ).)*?)</xref>#<xref ref-type="$1" rid="$2">$3$4</xref>#isg;
		#$Tmp=~s#<TLink href="([^"]*)">((?:(?!<\/TLink>|<Tlink ).)*?)<\/TLink>#<xref ref-type="" rid="$1">$2</xref>#isg;
		#$Tmp=~s#<xref ref-type="[^"]*" rid="(fig[^"]*)">#<xref ref-type="fig" rid="$1">#isg;
		#$Tmp=~s#<xref ref-type="[^"]*" rid="(tab[^"]*)">#<xref ref-type="table" rid="$1">#isg;
		#$Tmp=~s#<xref ref-type="[^"]*" rid="(box[^"]*)">#<xref ref-type="box" rid="$1">#isg;
                
                # usage, per body paragraph, with $num = current chapter number in scope:
                for my $type (qw(fig table box video casestudy exhibit)) {
                    $Tmp = ConvertLabel($Tmp, $type, $num);
                }

                $Tmp=~s{<label[^>]*>((?:(?!<\/label>).)*?)</label>}{

                        my $tag = $&;
                        $tag =~ s{<xref ref-type="[^"]+" rid="[^"]+">((?:(?!<\/xref>).)*?)</xref>}{$1}gs;
                        qq($tag);

                }gies;
                
		$Tmp=~s#<p([^>]*)>(\s*)\&lt;BX\&gt;(\s*)</p>#<box>#g;
		$Tmp=~s#<p([^>]*)>(\s*)\&lt;\/BX\&gt;(\s*)</p>#</box>#g;
                $Tmp=~s# class="FE-01-Title"# class="FE-01-BoxTitle"#;
		$Tmp=~s#<p([^>]*)>(\s*)(<bold>)?(\s*)\&lt;box\&gt;(\s*)(</bold>)?(\s*)</p>#<box>#g;
		$Tmp=~s#<p([^>]*)>(\s*)(<bold>)?(\s*)\&lt;/box\&gt;(\s*)(</bold>)?(\s*)</p>#</box>#g;
                
		#boxed text process
		$Tmp=~s{<box>((?:(?!<\/box>).)*?)</box>}{

			my $tag = $&;
			$tag =~ s{<p class="([^"]*)(BoxTitle|ExhibitCaption)">((?:(?!<\/p>).)*?)</p>}{
				my $tagg = $&;
				$tagg =~ s{<p class="([^"]*)(BoxTitle|ExhibitCaption)">((?:(?!<\/p>).)*?)</p>}{$3};
				$tagg =~ s{(<strong>|</strong>)}{}gi;
				$tagg =~ s{(<xref[^>]*>|</xref>)}{}gi;
                                if ($tagg =~ m{Box ([0-9]+)(\.)?([0-9]*)?})
                                {
                                        $tagg =~ s{Box ([0-9]+)(\.)?([0-9]*)?}{<boxed-text id="box$1_$3"><label>Box<space>$1$2$3</label><caption><title>};
                                }
                                elsif ($tagg =~ m{Exhibit ([0-9]+)(\.)?([0-9]*)?})
                                {
                                        $tagg =~ s{Exhibit ([0-9]+)(\.)?([0-9]*)?}{<boxed-text content-type="exhibit" id="exhibit$1_$3"><label>Exhibit<space>$1$2$3</label><caption><title>};
                                }
                                else
                                {
                                        $tagg = "<boxed-text><caption><title>" . $tagg;
                                }
				$tagg = $tagg . "</title></caption>";
				qq ($tagg);

			}gesi;
			$tag =~ s{(<box>|</box>)}{}g;
			$tag =~ s{<p class="Box-01-Para-FL">}{<p>}g;
			$tag =~ s{<p class="Exhibit-Para-FL">}{<p>}g;
                        $tag =~ s{<p class="Box-01-Head1"}{<p class="Head1"}g;
			$tag =~ s{<p class="Box-01-Note">((?:(?!<\/p>).)*?)</p>}{<attrib>$1</attrib>}g;
			$tag =~ s{<p class="Box-01-Source">((?:(?!<\/p>).)*?)</p>}{<attrib>$1</attrib>}g;
			$tag =~ s{<p class="Box-01-Equation">((?:(?!<\/p>).)*?)</p>}{<disp-formula>$1</disp-formula>}g;

                    $tag=~s{ class="ExhibitHeading}{ class="Head}g;
                    $tag=~s{<p class="H(?:ead)?(2|3|4|5|6)">((?:(?!<p |<\/p>).)*?)<\/p>}{"<sec$1 disp-level=\"level".($1-1)."\">\n<title>$2<\/title>\n<\/sec$1>"}gsie;
                    $tag=~s{<p class="H(?:ead)?(0|1|2|3|4|5|6)">((?:(?!<p |<\/p>).)*?)<\/p>}{<sec$1 disp-level="level$1">\n<title>$2<\/title>\n<\/sec$1>}gsi;
                    $tag=~s{^(.*?)$}{SecLevel("$&")}gsie;
                    $tag=~s{<p class="ExhibitNote">((?:(?!<\/p>).)*?)</p>}{<attrib>$1</attrib>}g;
                        $tag = $tag . "</boxed-text>";
			qq ($tag);

		}gesi;


		$Tmp =~ s{<space>}{ }ig;
		$Tmp =~ s{<p class="Exhibit-UN-TableBody">}{<p>}ig;
                $Tmp =~ s{<p class="BulletList1Source">((?:(?!</p>).)*?)</p>}{\n<p>$1</p>}g;
                $Tmp =~ s{<ext-link ext-link-type="uri" xlink:href="([^"]+)">(\s*)<ext-link ext-link-type="uri" xlink:href="([^"]+)">((?:(?!</ext-link>).)*?)</ext-link>(\s*)</ext-link>}{<ext-link ext-link-type="uri" xlink:href="$3">$4</ext-link>}g;
                $Tmp =~ s{<ext-link ext-link-type="doi" xlink:href="([^"]+)">(\s*)<ext-link ext-link-type="doi" xlink:href="([^"]+)">((?:(?!</ext-link>).)*?)</ext-link>(\s*)</ext-link>}{<ext-link ext-link-type="doi" xlink:href="$3">$4</ext-link>}g;
                $Tmp =~ s{(<cf21>|</cf21>)}{}g;
                $Tmp =~ s{<TLink href="([^"]+)">}{}g;
                $Tmp =~ s{</TLink>}{}g;
                $Tmp =~ s{(<bibetal>|</bibetal>)}{}g;
                $Tmp =~ s{(<FigureCitation>|<TableCitation>|<BoxCitation>|</FigureCitation>|</TableCitation>|</BoxCitation>|)}{}g;
		my $finalCont = "$booMeta$Tmp";

		#Final clean up
		$finalCont=~s#<p([^>]*)>((?:(?!<tab\/>).)*?)<tab\/>#<p$1>#ig;
		$finalCont=~s#<title>\&lt;LO\&gt;#<title>#isg;
		$finalCont=~s#<title>\&lt;SH1\&gt;#<title>#isg;
		$finalCont=~s#<p>\&lt;metadata\&gt;</p>##isg;
                $finalCont=~s#<strong>#<bold>#isg;
                $finalCont=~s#</strong>#</bold>#isg;
		$finalCont=~s#<p>\&lt;\/metadata\&gt;</p>##isg;
		$finalCont=~s#<p([^>]*)>(<bold>)?\&lt;online only\&gt;(<\/bold>)?</p>##isg;
		$finalCont=~s#<p([^>]*)>(<bold>)?\&lt;\/online only\&gt;(<\/bold>)?</p>##isg;
		$finalCont=~s#<p([^>]*)>(<bold>)?\&lt;e-only\&gt;(<\/bold>)?</p>##isg;
		$finalCont=~s#<p([^>]*)>(<bold>)?\&lt;\/e-only\&gt;(<\/bold>)?</p>##isg;
		#$finalCont=~s#<p([^>]*)>(<bold>)?\&lt;online only\&gt;(<\/bold>)?</p>#<p>&lt;onlineonly&gt;</p>#isg;
		#$finalCont=~s#<p([^>]*)>(<bold>)?\&lt;\/online only\&gt;(<\/bold>)?</p>#<p>&lt;/onlineonly&gt;</p>#isg;
		#$finalCont=~s# class="Reference-Alphabetical">#<p class="ReferenceAlphabetical">#isg;
		#$finalCont=~s#<p class="([^"]+)">#<p><\!-- $1 -->#isg;
		$finalCont=~s#\n\s*\n#\n#isg;

		# set floating elements into cited place
		my %float;
		
		# capture figures and tables
		while ($finalCont =~ m{(<(fig|boxed-text|table-wrap)\b[^>]*\bid="([^"]+)"[^>]*>.*?</\2>)}sgx) {
		    $float{$3} = $1;
		}

		# remove the block
		$finalCont =~ s{(<(fig|boxed-text|table-wrap)\b[^>]*\bid="([^"]+)"[^>]*>.*?</\2>)}{}sgx;
		
		foreach my $rid (keys %float) {
			$finalCont =~ s{(<p\b[^>]*>.*?<xref\b[^>]*\brid="$rid"[^>]*>.*?</p>)}{

			    my ($p) = ($1);

			    # If matching float exists
			    if (exists $float{$rid}) {

				my $block = $float{$rid};

				# Prevrent double insertion
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
			$finalCont =~ s{</ref-list>}{</ref-list>\n$float{$rid}}s;
		}
                
                while ($finalCont=~ m{\@q([0-9]+)\@})
                {
                        my $qnum = $1;
                        my $query;
                        if($finalCont=~ m{<query$qnum>((?:(?!<\/query).)*?)</query$qnum>})
                        {
                           $query = $&;
                           $finalCont=~ s{<query$qnum>((?:(?!<\/query).)*?)</query$qnum>}{}s;
                        }
                        $finalCont=~ s{\@q$qnum\@}{$query};
                }
                
                $finalCont=~ s{\@hi\@}{<!--<highlight>-->}g;
                $finalCont=~ s{\@\/hi\@}{<!--<\/highlight>-->}g;
                $finalCont=~ s{<query([0-9]+)>}{<!--<query>-->}g;
                $finalCont=~ s{</query([0-9]+)>}{<!--</query>-->}g;
                $finalCont=~ s{<abstract>(\s*)<title>Abstract</title>(\s*)\&abstract;(\s*)</abstract>}{}g;
                $finalCont=~ s{<kwd-group kwd-group-type="author">(\s*)<title>&KeyTermsHeading;</title>(\s*)&kwd;(\s*)</kwd-group>}{}g;
		$finalCont=~ s#(\n+)#\n#isg;
                $finalCont=~ s{  }{ }g;
                $finalCont=~ s{<p>(\s*)<\/p>}{}g;
                $finalCont=~ s{(<apple-converted-space>|</apple-converted-space>)}{}g;
                $finalCont=~ s{<p>\&lt;onlineonly\&gt;</p>(\s*)</([^>]*)>}{</$2>$1<p>\&lt;onlineonly\&gt;</p>}gs;
                $finalCont=~ s{<p>\&lt;\/onlineonly\&gt;</p>(\s*)</([^>]*)>}{</$2>$1<p>\&lt;\/onlineonly\&gt;</p>}gs;
                $finalCont=~ s{<contrib-group>(\s*)&contrib;(\s*)</contrib-group>}{}g;
                $finalCont=~ s{&(?:num|ChLabel|ChTitle|contrib|abstract|kwd|KeyTermsHeading);}{}g;
                #$finalCont=~ s{(<nocitebib>|</nocitebib>)}{}g;
		$finalCont=~ s{&(?!amp;|lt;|gt;|quot;|apos;|#[0-9]+;|#x[0-9a-fA-F]+;)}{&#x0026;}isg;
		$finalCont=~ s{([^\x00-\x7F])}{sprintf('&#x%04X;', ord($1))}ge;
		&WriteFile("$Final_File", "$finalCont", "HTML");
		system($PYTHON_BIN, "$File_Path/utf8_converter.py", "$Final_File");

		&DTDvalidate("$Final_File");
#========================= Sub Functions =========================#

sub DTDvalidate
{
	my $xml_file = shift;

	# -------- Log file (same name + .log) --------
	my ($name, $path, $suffix) = fileparse($xml_file, qr/\.[^.]*/);
	my $log_file = $path . $name . ".log";

	open(my $LOG, '>', $log_file) or die "Cannot open log file: $!";

	print $LOG "BITS DTD Validation Log\n";
	print $LOG "Input File : $xml_file\n";
	print $LOG "---------------------------------\n";

	# -------- XML Parser --------
	my $parser = XML::LibXML->new(
	    load_ext_dtd => 1,
	    validation   => 1
	);

	eval {
	    $parser->parse_file($xml_file);
	};

	if ($@) {
	    print $LOG "? VALIDATION FAILED\n\n";
	    print $LOG "$@\n";
	    print "Validation FAILED. See log: $log_file\n";
	} else {
	    print $LOG "? VALIDATION PASSED\n";
	    print "Validation PASSED.\n";
	}

	close $LOG;
}
		
sub ReadFile
{
	my ($infile, $type)=@_;
	open (IN,"<$infile") or die "Unable to open $type file $infile: $!";
	undef $/; my $cont=<IN>;
	close IN;
	return $cont;
}
sub WriteFile
{
	my $outfile=shift;
	my $cont=shift;
	my $type=shift;
	open (OUT,">$outfile") or die "Unable to write $type file $outfile: $!";
	print OUT $cont;
	close OUT;
}
sub listHead{
	my $tmp = shift;
	my $type = shift;
	if($type=~m#(BulletList1first|LearnObj-BulletList1-first)#is){
	#		$tmp=~s#<p class="(BulletList1first[0-9]+|LearnObj-BulletList1-first)">((?:(?!<\/p>|<p ).)*?)<\/p>#<list list-type="bullet">\n<list-item>\n<p>$2</p>\n</list-item><\/bl1>#isg;
	}else{
	#		$tmp=~s#<p class="(BulletList1|LearnObj-BulletList1)">((?:(?!<\/p>|<p ).)*?)<\/p>#<bl1>\n<list-item>\n<p>$2</p>\n</list-item><\/bl1>#isg;
		$tmp=~s#<p class="(KeyTerm)">((?:(?!<\/p>|<p ).)*?)<\/p>#<kt1>\n<list-item>\n<p id="term&seq;">$2</p>\n</list-item><\/kt1>#isg;
	}
	return $tmp;
}

sub SecLevel {
    my $lvlCont = shift;

    # 1. Strip existing closing section tags
    $lvlCont =~ s#</sec\d+>##gsi;

    # 2. Append a dummy marker before </body> so open sections close at the end
    $lvlCont =~ s#<\/body>#<sec_end><\/body>#is;

    # 3. Mark start of section tags for splitting
    $lvlCont =~ s#(<sec\d+|<sec_end>)#<enter>$1#gsi;

    my @body = split('<enter>', $lvlCont);
    my @stack; # Stack to keep track of active section levels

    foreach my $line (@body) {
        if ($line =~ /<sec(\d+)[^\>]*?>/) {
            my $current_lvl = $1; # Pure integer (e.g., 0, 1, 2)

            my $closetags = '';
            # Close any sections in the stack that are deeper or equal to current level
            while (@stack && $stack[-1] >= $current_lvl) {
                my $closed_lvl = pop @stack;
                $closetags .= "</sec$closed_lvl>\n";
            }

            $line = $closetags . $line;
            push @stack, $current_lvl;
        }
        elsif ($line =~ /<sec_end>/) {
            # Close all remaining open sections at the end of the document
            my $closetags = '';
            while (@stack) {
                my $closed_lvl = pop @stack;
                $closetags .= "</sec$closed_lvl>\n";
            }
            $line =~ s/<sec_end>/$closetags/;
        }
    }

    my $result = join("", @body);

    # Cleanup excess line breaks
    $result =~ s#\n{2,}#\n#gsi;
    $result =~ s#<(\/)?sec(\d+)#<$1sec#gsi;

    return $result;
}

sub SecLevel_dummy{
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
    $text=~s#<(bibedfname|bibedsurname)>((?:(?!<\/\1>|<\1>).)*?)<\/\1>#<pge><strname>$&</strname></pge>#isg;
    $text=~s#</pg>([^<>]*)<pg>#$1#isg;
    $text=~s#<\/bibedfname></strname>([^<>a-z]*)<strname><bibedsurname>#<\/bibedfname>$1<bibedsurname>#isg;
    $text=~s#<\/bibsurname></strname>([^<>a-z]*)<strname><bibfname>#<\/bibsurname>$1<bibfname>#isg;
    $text=~s#<(\/)?bibsurname>#<$1surname>#isg;
    $text=~s#<(\/)?bibfname>#<$1given-names>#isg;
    $text=~s#<(\/)?bibedsurname>#<$1surname>#isg;
    $text=~s#<(\/)?bibedfname>#<$1given-names>#isg;
    $text=~s#<(\/)?strname>#<$1string-name>#isg;
    $text=~s#<pg>#<person-group person-group-type="author">#isg;
    $text=~s#<\/pg>#<\/person-group>#isg;
    $text=~s#<pge>#<person-group person-group-type="editor">#isg;
    $text=~s#<\/pge>#<\/person-group>#isg;
    $text=~s#<\/string-name><\/person-group> <person-group person-group-type="editor"><string-name># #ig;
    $text=~s#<\/person-group> \&\#x0026; <person-group person-group-type="editor"># #g;
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
    
    
    if($text=~m#(<\/issue>|<\/article-title>)#is){
	    $text=~s#<p class="Reference-Alphabetical">#<ref id="bid_${num}_&seq2;"><mixed-citation publication-type="article">#isg;
	    $text=~s#<p class="Reference-Numbered">#<ref id="bid_${num}_&seq2;"><mixed-citation publication-type="article">#isg;
    }elsif($text=~m#(<\/chapter-title>|<\/publisher-name>|<\/publisher-loc>)#is){
	    $text=~s#<p class="Reference-Alphabetical">#<ref id="bid_${num}_&seq2;"><mixed-citation publication-type="book">#isg;
	    $text=~s#<p class="Reference-Numbered">#<ref id="bid_${num}_&seq2;"><mixed-citation publication-type="book">#isg;
    }if($text=~m#(<\/ext-link>)#is){
	    $text=~s#<p class="Reference-Alphabetical">#<ref id="bid_${num}_&seq2;"><mixed-citation publication-type="web">#isg;
	    $text=~s#<p class="Reference-Numbered">#<ref id="bid_${num}_&seq2;"><mixed-citation publication-type="web">#isg;
    }else{
	    $text=~s#<p class="Reference-Alphabetical">#<ref id="bid_${num}_&seq2;"><mixed-citation publication-type="other">#isg;
	    $text=~s#<p class="Reference-Numbered">#<ref id="bid_${num}_&seq2;"><mixed-citation publication-type="other">#isg;
    }
    $text=~s#<\/p>#<\/mixed-citation><\/ref>#isg;

    $text=~s#</collab><collab>##isg;
    $text=~s#</bibbook><bibbook>##isg;
    $text=~s#</bibeditionno><bibeditionno>##isg;
    $text=~s#<nocitebib>##isg;
    $text=~s#</nocitebib>##isg;
    $text=~s#<bibbook>#<source>#isg;
    $text=~s#</bibbook>#</source>#isg;
    $text=~s#<bibeditionno>#<edition>#isg;
    $text=~s#</bibeditionno>#</edition>#isg;
    $text=~s#<bibinstitution>#<institution>#isg;
    $text=~s#</bibinstitution>#</institution>#isg;
    $text=~s#<bibnumber>#<label>#isg;
    $text=~s#</bibnumber>#</label>#isg;
    
    $text=~s#<ref id="([^"]+)"><mixed-citation publication-type="([^"]+)"><label>([^>]+)</label>#<ref id="bib_${num}_$3"><mixed-citation publication-type="$2"><label>$3</label>#isg;
    $text=~s#</label>\.(\s*)#\.</label>#isg;
    
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

sub LearnObjectives
{
	my ($Title,$Intro,$Items,$ChNum)=@_;

	$Title=~s{<[^>]+>}{}gsi;      # <title> already implies bold - strip inline tags e.g. <strong>

	my @List;
	while($Items=~m{<p class=\"LearnObj-NumberList1(?:first|last)?0?\">((?:(?!<\/p>).)*?)<\/p>}gsi)
	{
		push(@List,"<list-item><p>$1</p></list-item>");
	}
	while($Items=~m{<p class=\"LearnObj-BulletList1(?:first|last)?0?\">((?:(?!<\/p>).)*?)<\/p>}gsi)
	{
		push(@List,"<list-item><p>$1</p></list-item>");
	}
	while($Items=~m{<p class=\"LearnObj-UL-FL1(?:first|last)?0?\">((?:(?!<\/p>).)*?)<\/p>}gsi)
	{
		push(@List,"<list-item><p>$1</p></list-item>");
	}
	my $ListItems=join("\n",@List);

	my $IntroPara = (defined $Intro && $Intro ne "") ? "<p>$Intro</p>\n" : "";
        
        if ($Items=~m{<p class=\"LearnObj-Bullet})
        {
                return "<sec disp-level=\"LearnObject\" id=\"ch${ChNum}lev1sec&seq1;\"><title>$Title</title>\n${IntroPara}\n<list list-type=\"bullet\">\n$ListItems\n</list>\n</sec>";
        }
        elsif ($Items=~m{<p class=\"LearnObj-Number})
        {
                return "<sec disp-level=\"LearnObject\" id=\"ch${ChNum}lev1sec&seq1;\"><title>$Title</title>\n${IntroPara}\n<list list-type=\"order\">\n$ListItems\n</list>\n</sec>";
        }
        else
        {
                return "<sec disp-level=\"LearnObject\" id=\"ch${ChNum}lev1sec&seq1;\"><title>$Title</title>\n${IntroPara}\n<list list-type=\"none\">\n$ListItems\n</list>\n</sec>";
        }
}

sub CaseStudyBulletList
{
	my ($Items,$ChNum)=@_;

	my @List;
	while($Items=~m{<p class=\"CaseStudy-BulletList1(?:first|last)?0?\">((?:(?!<\/p>).)*?)<\/p>}gsi)
	{
		push(@List,"<list-item><p>$1</p></list-item>");
	}
	while($Items=~m{<p class=\"BulletList1[_]?(?:first|last)?0?\">((?:(?!<\/p>).)*?)<\/p>}gsi)
	{
		push(@List,"<list-item><p>$1</p></list-item>");
	}
	my $ListItems=join("\n",@List);

	return "\n<list list-type=\"bullet\">\n$ListItems\n</list>\n";
}

sub NestLists {
    my ($block) = @_;
    my @items;

    # Extract list type prefix (RL/BL), level number (1/2), and inner content
    while ($block =~ m{<(RL|BL)([1-9])>(.*?)</\1\2>}gsi) {
        push @items, { type => $1, level => $2, text => $3 };
    }

    my $out = "";
    my @stack;

    foreach my $item (@items) {
        my $type  = $item->{type};
        my $level = $item->{level};
        my $text  = $item->{text};

        my $list_type = ($type eq 'RL') ? 'roman-upper' : 'bullet';

        # Step down: close child lists and child list-items
        while (@stack && $stack[-1]{level} > $level) {
            $out .= "\n</list>\n</list-item>";
            pop @stack;
        }

        # Same level: close preceding list-item
        if (@stack && $stack[-1]{level} == $level) {
            $out .= "</list-item>\n";
        }
        # Step up: open sub-list inside parent <list-item>
        elsif (!@stack || $stack[-1]{level} < $level) {
            $out .= "\n<list list-type=\"$list_type\">\n";
            push @stack, { level => $level, type => $type };
        }

        $out .= "<list-item><p>$text</p>";
    }

    # Close remaining open tags on stack
    while (@stack) {
        $out .= "\n</list>\n</list-item>";
        pop @stack;
    }

    # Remove extra closing list-item from the outermost level
    $out =~ s{\n</list-item>$}{};

    return $out;
}

sub BulletList
{
	my ($Items,$ChNum)=@_;

	my @List;
	while($Items=~m{<p class=\"BulletList2[_]?(?:first|last)?0?\">((?:(?!<\/p>).)*?)<\/p>}gsi)
	{
		push(@List,"<list-item><p>$1</p></list-item>");
	}
	while($Items=~m{<p class=\"BulletList1[_]?(?:first|last)?0?\">((?:(?!<\/p>).)*?)<\/p>}gsi)
	{
		push(@List,"<list-item><p>$1</p></list-item>");
	}
	my $ListItems=join("\n",@List);
        if ($Items=~m{<p class=\"BulletList2})
        {
                return "\n<list list-type=\"bullet2\">\n$ListItems\n</list>\n";
        }
        else
        {
                return "\n<list list-type=\"bullet\">\n$ListItems\n</list>\n";
        }
}


sub CaseStudyNumberList
{
	my ($Items,$ChNum)=@_;

	my @List;
	while($Items=~m{<p class=\"CaseStudy-NumberList1(?:first|last)?0?\">((?:(?!<\/p>).)*?)<\/p>}gsi)
	{
		push(@List,"<list-item><p>$1</p></list-item>");
	}
	while($Items=~m{<p class=\"NumberList1[_]?(?:first|last)?0?\">((?:(?!<\/p>).)*?)<\/p>}gsi)
	{
		push(@List,"<list-item><p>$1</p></list-item>");
	}
	my $ListItems=join("\n",@List);

	return "\n<list list-type=\"order\">\n$ListItems\n</list>\n";
}

sub CaseStudyUnNumberList
{
	my ($Items,$ChNum)=@_;

	my @List;
	while($Items=~m{<p class=\"CaseStudy-UL-FL1(?:first|last)?0?\">((?:(?!<\/p>).)*?)<\/p>}gsi)
	{
		push(@List,"<list-item><p>$1</p></list-item>");
	}
	while($Items=~m{<p class=\"Exhibit-UL-FL1(?:first|last)?0?\">((?:(?!<\/p>).)*?)<\/p>}gsi)
	{
		push(@List,"<list-item><p>$1</p></list-item>");
	}
	while($Items=~m{<p class=\"UL-FL1[_]?(?:first|last)?0?\">((?:(?!<\/p>).)*?)<\/p>}gsi)
	{
		push(@List,"<list-item><p>$1</p></list-item>");
	}
	my $ListItems=join("\n",@List);

	return "\n<list list-type=\"none\">\n$ListItems\n</list>\n";
}

sub GlossaryUnNumberList {
    my ($Items) = @_;
    my @List;

    while ($Items =~ m{<p class=\"GlossaryTermDefinitionUL-FL1[_]?(?:first|last)?0?\">((?:(?!<\/p>).)*?)<\/p>}gsi) {
        push(@List, "<list-item><p>$1</p></list-item>");
    }

    my $ListItems = join("\n", @List);
    return "\n<list list-type=\"none\">\n$ListItems\n</list>\n";
}


sub caseStudy{
    my $text = shift;
    $text=~s# class="FE-# class="#isg;
    #$text=~s# class="CaseStudy-# class="CaseStudy#isg;
    $text=~s# class="CaseStudyHeading# class="Head#isg;
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
    $text=~s#<tgroup(\s*)[^<>]*\/>##gsi;
    $text=~s#<colspec [^<>]*colnum="([^"]+)"[^<>]*\/>#<colgroup>\n<col content-type="$1"/>\n</colgroup>#gsi;
    $text=~s#<\/colgroup>\s*<colgroup>#\n#gsi;
    $text=~s#<entry([^<>]*)( colname="[^"]*")([^<>]*)>#<td$1$3>#gsi;
    $text=~s#<\/entry>#<\/td>#gsi;
    $text=~s#<row([^<>]*)>#<tr>#gsi;
    $text=~s#<\/row>#<\/tr>#gsi;
    #$text=~s#<p class="([^"]+)">#<p>#gsi;
    $text=~s#<tbody>((?:(?!<tbody |<\/tbody>).)*?TableColumnHead(?:(?!<tbody |<\/tbody>).)*?)<\/tbody>#tablethead($&)#gsie;
    $text=~s#<thead>((?:(?!<thead |<\/thead>).)*?TableColumnHead(?:(?!<thead |<\/thead>).)*?)<\/thead>#tablethead($&)#gsie;
    $text=~s##<p>#isg;
    #$text=~s#<\/p>\s*<p class="TableBody[0-9]*">#<br\/>#isg;
    $text=~s#<\/p>\s*<p class="TableBody[0-9]*">#<\/p>\n<p>#isg;
    $text=~s#<p class="TableBody[0-9]*">#<p>#isg;
    #$text=~s#(<\/p>|<p(?: [^<>]*)?>)##isg;
    $text=~s#\s+<\/t(h|d)>#<\/t$1>#isg;
    $text=~s#</tbody>\s*<tbody>##isg;
    $text=~s{(<fn>)((?:(?!<fn |<\/fn>).)*?)(</fn>)</table>}{</table>\n<table-wrap-foot>$1$2$3</table-wrap-foot>}g;
    $text=~s{ colsep="(0|1)"}{}g;
    $text=~s{ rowsep="(0|1)"}{}g;
    $text=~s{<fn>}{<fn><p>}g;
    $text=~s{</fn>}{</p></fn>}g;
    $text=~s{<p><p>}{<p>}g;
    $text=~s{</p></p>}{</p>}g;
    $text=~s{ align="both"}{ align="left"}g;
    $text=~s{<list-item><p>\&\#x2022; }{<list-item><p>}g;
    return "<table-wrap>$text</table-wrap>";
}

sub tablethead{
    my $text = shift;
    $text=~s#<td([^>]*)>#<th$1>#gsi;
    $text=~s#<\/td>#<\/th>#gsi;
    $text=~s#<tbody>((?:(?!<tbody |<\/tbody>).)*?TableColumnHead(?:(?!<tbody |<\/tbody>).)*?)<\/tbody>#<thead>$1<\/thead>#gsi;
    $text=~s#<p class="([^"]*)TableColumnHead[0-9]*">#<p>#isg;
    return $text;
}

sub build_regex
{
    my ($labelRe) = @_;
    return qr/\b($labelRe)\.?\s+(${ITEM}(?:${CONNECT}${ITEM})*)\b/;
}

sub Xref
{
    my ($chNum, $type, $matchedLabel, $matchedNums) = @_;
    my $info = $LabelMap->{$type};

    my @items = ($matchedNums =~ m/(${ITEM})/g);
    my @rids  = map {
        my $t = $_;
        $t = "${chNum}.${t}" unless $t =~ /[.\x{2011}]/;   # abbreviated tail -> prepend chapter
        "$info->{prefix}$t";
    } @items;

    my $ridAttr = join(' ', @rids);
    $ridAttr=~ s{\.}{\_};
    return "<xref ref-type=\"$info->{reftype}\" rid=\"$ridAttr\">$matchedLabel $matchedNums</xref>";
}

sub ConvertLabel
{
    my ($text, $type, $chNum) = @_;
    my $re = build_regex($LabelMap->{$type}{re});
    $text =~ s/$re/&Xref($chNum,$type,$1,$2)/ge;
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
	$tmp=~ s#\-#\\\-#isg;
	$tmp=~ s#\[#\\\[#isg;
	$tmp=~ s#\]#\\\]#isg;
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
	$refCall =~ s{&(?!amp;|lt;|gt;|quot;|apos;|#[0-9]+;|#x[0-9a-fA-F]+;)}{&#x0026;}isg;
	return $refCall;
}

sub resolve_nocitebib {
    my $content = shift;

    my $by_id = build_ref_index($content);

    my @unresolved;
    $content =~ s{<nocitebib>((?:(?!</nocitebib>).)*)</nocitebib>}{
        my $inner = $1;
        if ($inner =~ m{<sup>})
        {
                $inner =~ s{([0-9]+)}{<xref ref-type="bibr" rid="bib_${num}_$1">$1</xref>}g;
                qq{$inner};
        }
        else
        {
                my $id    = find_ref_id($inner, $by_id);
                my $inner_xml = $inner;
                $inner_xml =~ s{&(?!amp;|lt;|gt;|quot;|apos;|#[0-9]+;|#x[0-9a-fA-F]+;)}{&#x0026;}isg;
                if (defined $id) {
                    qq{<xref ref-type="bibr" rid="$id">$inner_xml</xref>};
                } else {
                    push @unresolved, $inner;
                    "<nocitebib>$inner_xml</nocitebib>";   # leave for manual review
                }
        }
    }gesx;

    if (@unresolved) {
        warn "Unresolved citations (" . scalar(@unresolved) . "):\n"
           . join("\n", map { "  - $_" } @unresolved) . "\n";
    }
    return $content;
}

# ---- Build a lookup of ref id => { first surname/org, all surnames, acronym, year } ----
sub build_ref_index {
    my $content = shift;
    my %by_id;

    while ($content =~ m{<ref\s+id="([^"]+)"[^>]*>(.*?)</ref>}gs) {
        my ($id, $body) = ($1, $2);

        my @surnames = ($body =~ m{<surname>([^<]*)</surname>}g);
        my ($collab) = ($body =~ m{<collab>([^<]*)</collab>});
        my ($year)   = ($body =~ m{<year>(\d+)});

        my $first = @surnames ? $surnames[0] : ($collab // '');

        $by_id{$id} = {
            first    => norm($first),
            surnames => [ map { norm($_) } @surnames ],
            acronym  => $collab ? norm(acronym_of($collab)) : '',
            year     => $year // '',
        };
    }
    return \%by_id;
}

sub acronym_of {
    my $collab = shift;
    my @words = grep { $_ !~ /^(?:of|the|for|and|on|in|a|an)$/i }
                split /\s+/, $collab;
    return join('', map { substr($_, 0, 1) } @words);
}

sub norm {
    my $s = shift // '';
    $s = lc $s;
    $s =~ s/&#x[0-9a-f]+;//gi;   # drop stray numeric-entity remnants
    $s =~ s/[^a-z]//g;          # letters only - ignores spaces/accents/punct
    return $s;
}

# ---- Match one in-text citation string against the ref index ----
sub find_ref_id {
    my ($text, $by_id) = @_;

    (my $t = $text) =~ s/&#x0026;|&amp;/&/gi;   # normalize "&" entity for parsing

    my @yrs = ($t =~ /(\d{4}[a-z]?)/g);
    return undef unless @yrs;
    my $year_full = $yrs[-1];                    # last year-like token wins
    (my $year = $year_full) =~ s/[a-z]$//;

    my $authors = $t;
    $authors =~ s/\Q$year_full\E.*$//s;
    my ($acronym) = $t =~ /\[([^\]]+)\]/;
    $authors =~ s/\[[^\]]*\]//g;

    my $et_al = ($authors =~ s/\bet\s*al\.?\s*,?\s*$//i);
    $authors =~ s/,\s*$//;

    my @names = grep { length }
                map  { s/^\s+|\s+$//gr }
                split /\s*(?:&|,?\s+and\s+)\s*/i, $authors;

    my @keys = map { norm($_) } @names;
    push @keys, norm($acronym) if $acronym;

    for my $id (keys %$by_id) {
        my $ref = $by_id->{$id};
        next unless $ref->{year} eq $year;

        if ($et_al || @keys == 1) {
            return $id if grep { $_ eq $ref->{first} || $_ eq $ref->{acronym} } @keys;
        }
        else {
            my $all_match = 1;
            for my $k (@keys) {
                $all_match = 0 unless grep { $_ eq $k } @{ $ref->{surnames} };
            }
            return $id if $all_match;
        }
    }
    return undef;
}

#</citebib>
#<volume><italic>24</italic></volume>
#ext-link-type="doi"
}
