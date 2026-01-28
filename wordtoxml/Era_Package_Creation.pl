use strict;
use Archive::Zip qw/ :ERROR_CODES :CONSTANTS /;
use File::Basename;
use Cwd 'abs_path';
use Config::IniFiles;
use File::Copy;
use File::Copy::Recursive qw(dircopy);
use File::Find;
use Uniq;
use utf8;

my $InPath=$ARGV[0];
my $InputINI=$ARGV[1];

my $ExePath=abs_path($0);
my ($ExeName, $ExeDir, $ExeSuffix) = fileparse($ExePath, "\.(pl|exe)");

opendir(DIR,$InPath) or die("Unable to open html folder $InPath: $!");
my @htmlfiles=grep{/\.(x?html?)$/i} readdir(DIR); closedir DIR;
# @htmlfiles = sort (@htmlfiles);
@htmlfiles = sort {$a <=> $b} @htmlfiles;

my @images;
if (-d "$InPath/img")
{
	opendir(DIR,"$InPath/img") or die "Unable to open $InPath/img";
	@images=grep{/(?:\.jpg|\.jpeg|\.png|\.gif|\.tif|\.ico|\.eps|\.bmp)/i} readdir(DIR);
	closedir DIR;
}
	
#--- Content File ---#

	opendir(DIR,$InPath) or die("Unable to open html folder $InPath: $!");
	my @TOCfiles=grep{/(\_|\b)(toc|contents?)\.(x?html?)$/i} readdir(DIR); closedir DIR;
	@TOCfiles = sort (@TOCfiles);

#--- Temp

	# my $ini=&ReadFileDec('E:\Ram\E-Books\Non-RTS_E-Books\Tamil_OCR\Perl_Source\Title_Map.ini', "");
	# &WriteFile('E:\Ram\E-Books\Non-RTS_E-Books\Tamil_OCR\Perl_Source\Title_Map.ini',"$ini","");

#-------------------------------- INPUT INI ------------------------------#

	my ($ComboCoverInput, $Begin_Chapter, $Book_TitleOne, $Book_Title, $ISBN, $Pages, $Publisher, $Language, $Creator, $CreatorOne, $Version);

	my $INICONT=&ReadFileDec("$InputINI", "INI");
	&WriteFile("$InputINI", "$INICONT", "INI");
	
	if (-f "$InputINI")
	{
			my $cfg = new Config::IniFiles(-file => "$InputINI");
			$ComboCoverInput=$cfg->val('MAIN', 'Cover');
			$Begin_Chapter=$cfg->val('MAIN', 'Begin_Chapter');
			$Book_Title=$cfg->val('MAIN', 'Book_Title');
			$Book_TitleOne=$cfg->val('MAIN', 'Book_Title1');
			$ISBN=$cfg->val('MAIN', 'ISBN');
			$Pages=$cfg->val('MAIN', 'Pages');
			$Publisher=$cfg->val('MAIN', 'Publisher');
			$Language=$cfg->val('MAIN', 'Language');
			$Language=uc($Language);
			$Creator=$cfg->val('MAIN', 'Creator');
			$CreatorOne=$cfg->val('MAIN', 'Creator1');
			$Version=$cfg->val('MAIN', 'Version');
	}

	
#--------------------------- Title Map INI File --------------------------#

	my $ExeINI="${ExeDir}${ExeName}.ini";
	my ($TitleMapCont, $html_meta, $opf_cont, $ncx_cont, $nav_cont, $container_cont, $apple_cont, $Lang_cont);

	if (-f "$ExeINI")
	{
			my $cfg = new Config::IniFiles(-file => "$ExeINI");
			if ($Language ne '')
			{
					$TitleMapCont=$cfg->val('MAIN',"$Language");
			}
			$html_meta=$cfg->val('MAIN','HTML_META');
			$opf_cont=$cfg->val('MAIN','OPF_CONT');
			$ncx_cont=$cfg->val('MAIN','NCX_CONT');
			$nav_cont=$cfg->val('MAIN','NAV_CONT');
			$container_cont=$cfg->val('MAIN','CONTAINER');
			$apple_cont=$cfg->val('MAIN','APPLE');
			$Lang_cont=$cfg->val('MAIN','LANGUAGE');
	}

	my $Lang_code=$1 if ($Lang_cont=~m#$Language\=([^\n]*?)$#mi);
	
	
#--------------------------- Folder Creation ------------------------------#

	my $EpubFolder;
	if ($ISBN ne '')
	{
		$EpubFolder="$InPath/$ISBN";
	}
	else
	{
		$EpubFolder="$InPath/OUTPUT";
	}
	
	my $META_INF="$EpubFolder\\META-INF";
	my $OEBPS="$EpubFolder\\OEBPS";
	
	my $HtmlPath="$OEBPS\\html";
	my $CSSPath="$OEBPS\\css";
	my $IMGPath="$OEBPS\\img";
	
	foreach ("$EpubFolder", "$META_INF", "$OEBPS", "$HtmlPath", "$CSSPath", "$IMGPath")
	{
		mkdir("$_");
	}
	#-- version 3.0
	my $NAVPath="$OEBPS\\nav";
	if ($Version eq '3.0')
	{
		mkdir("$NAVPath");
	}
	
	#-- Create mimetype
	&WriteFile("$EpubFolder/mimetype", "application\/epub\+zip", "mimetype");
	
	#-- Create Container
	&WriteFile("$META_INF/container.xml", "$container_cont", "container");
	
	#-- Create apple ibooks xml
	&WriteFile("$META_INF/com.apple.ibooks.display-options.xml", "$apple_cont", "ibooks");

	#-- Copy CSS
	if (-f "$InPath/epub.css")
	{
		copy("${InPath}/epub.css", "$CSSPath/epub.css");
	}
	else
	{
		copy("${ExeDir}epub.css", "$CSSPath/epub.css");
	}
	
	#-- Css Updations
	my $css_cont=&ReadFile("$CSSPath/epub.css", "CSS");
	my $new_css;
	$new_css=$css_cont;
	$new_css.="\n\@media amzn-mobi\n\{\n".$css_cont."\n\}\n";
	$new_css.="\n\@media amzn-kf8\n\{\n".$css_cont."\n\}\n";
	&WriteFile("$CSSPath/epub.css", "$new_css", "CSS");
	
	#----------------------------- Cover ------------------------------#
	
	my $ComboCover;
	my $coverimagepath;
	if ($ComboCoverInput=~m#(?:\.jpg|\.jpeg)#is)
	{
			$coverimagepath=$ComboCoverInput;
			if (!-f "$coverimagepath")
			{
				die "Please enter absolute cover image path\n";
				exit;
			}
			if ($coverimagepath!~m#(?:\.jpg|\.jpeg)$#is)
			{
				die "Please enter correct image name\n";
				exit;
			}
	}
	elsif($ComboCoverInput=~m#(?:\.html|\.htm|\.xhtml|\.xhtm)#is)
	{
			$ComboCover=$ComboCoverInput;
	}
	
#--------------------------- OPF Creation --------------------------#

print "\n\n\tOPF creation\n";

	$opf_cont=~s/unique\-identifier\=\"[^\"]*?\"/unique-identifier="p$ISBN"/is;
	$opf_cont=~s/<dc:title>[^>]*?<\/dc:title>/<dc:title>$Book_Title<\/dc:title>/is;
	$opf_cont=~s/<dc\:identifier\s*id\=\"[^\"]*?\">[^>]*?<\/dc\:identifier>/<dc:identifier id="p$ISBN">$ISBN<\/dc:identifier>/is;
	$opf_cont=~s/<dc:format>[^>]*?<\/dc:format>/<dc:format>$Pages pages<\/dc:format>/is;
	$opf_cont=~s/<dc:publisher>[^>]*?<\/dc:publisher>/<dc:publisher>$Publisher<\/dc:publisher>/is;
	$opf_cont=~s/<dc:source>[^>]*?<\/dc:source>/<dc:source>$ISBN<\/dc:source>/is;
	$opf_cont=~s/<dc:language>[^>]*?<\/dc:language>/<dc:language>$Lang_code<\/dc:language>/is;
	$opf_cont=~s/(<dc:creator[^\>]*?>)[^>]*?<\/dc:creator>/$1$Creator<\/dc:creator>/is;
	$opf_cont=~s/(<meta refines="#creator"[^\>]*?>)[^>]*?<\/meta>/$1$CreatorOne<\/meta>/is;
	$opf_cont=~s/(<meta refines="#main-title"[^\>]*?>)[^>]*?<\/meta>/$1$Book_TitleOne<\/meta>/is;
	$opf_cont=~s/ xml:lang=""/ xml:lang="$Lang_code"/gsi;
	
	#-- Add HTML list
	my $html_count=1;
	foreach (@htmlfiles)
	{
		my $htmlname=$_;
		$opf_cont.="\n".'<item id="nav_'.$html_count.'" href="html/'.$htmlname.'" media-type="application/xhtml+xml"/>';
		$html_count++;
	}
	#-- Add Image List
	my $imag_count=1;
	foreach(@images)
	{
		my $imgname=$_;
		$imgname=~s/\.gif/\.jpg/is;
		$opf_cont.="\n".'<item id="img'.$imag_count.'" href="img/'.$imgname.'" media-type="image/jpeg"/>';
		$imag_count++;
	}

	
	# $opf_cont.="\n</manifest>\n<spine toc=\"ncx\">\n<itemref idref=\"cover\" linear=\"yes\"/>\n";
	$opf_cont.="\n</manifest>\n<spine toc=\"ncx\">\n";
	
	#-- Spine Section
	my $itemref_count=1;
	foreach(@htmlfiles)
	{
		$opf_cont.="<itemref idref=\"nav_${itemref_count}\"/>\n";
		$itemref_count++;
	}
	
	#-- Guide Section
	$opf_cont.="</spine>\n<guide>";
	$opf_cont.="\n<reference type=\"cover\" title=\"Cover\" href=\"html/$ComboCover\"/>";
	$opf_cont.="\n<reference type=\"text\" title=\"Begin reading\" href=\"html/$Begin_Chapter\"/>";
	$opf_cont.="\n<reference type=\"toc\" title=\"Table of Contents\" href=\"html/$TOCfiles[0]\"/>";
	$opf_cont.="\n</guide>\n</package>";

	&WriteFile_DecToChar("$OEBPS/content.opf", "$opf_cont", "OPF");


#--------------------------- NCX Creation --------------------------#
	
	my $toc_cont; #-- For TOC file
	my $nav_toc;  #-- For Nav Toc Content
	
print "\n\tNCX creation\n";

	$ncx_cont.="\n";
	$ncx_cont=~s#xml:lang=\"\"#xml:lang="$Lang_code"#is;
	$ncx_cont=~s#<meta\s*name\=\"dtb\:uid\"\s*content\=\"[^\"]*?\"\s*\/\s*>#<meta name="dtb:uid" content="$ISBN"\/>#is;
	$ncx_cont=~s#<meta\s*name\=\"dtb\:totalPageCount\"\s*content\=\"[^\"]*?\"\s*\/\s*>#<meta name="dtb:totalPageCount" content="$Pages"\/>#is;
	$ncx_cont=~s#<meta\s*name\=\"dtb\:maxPageNumber\"\s*content\=\"[^\"]*?\"\s*\/\s*>#<meta name="dtb:maxPageNumber" content="$Pages"\/>#is;
	$ncx_cont=~s#<docTitle>\s*<text>[^>]*?<\/text>\s*<\/docTitle>#<docTitle><text>$Book_Title<\/text><\/docTitle>#is;
	$ncx_cont=~s#<docAuthor><text>[^>]*?<\/text><\/docAuthor>#<docAuthor><text>$Creator<\/text><\/docAuthor>#is;

	foreach my $htmlname (@htmlfiles)
	{
			my $cont=&ReadFileDec("$InPath/$htmlname", "HTML");
			
			my $filename=$htmlname;
			$filename=~s#^\d+\_##gsi;
			$filename=~s#\.x?html?$##gsi;
			$filename=~s#\_Page$##gsi;
			
			my $link_cont;
			if ($TitleMapCont=~m#(?:\_|\b)$filename\=([^\n]*?)$#mi)
			{
				$link_cont=$1;
			}
			elsif ($cont=~m#<body[^\>]*?>\s*<(\w+)[^\>]*?>(.*?)</\1>#si)
			{
					$link_cont=$2;
					$link_cont=~s#<[^\>]*?>##gsi;
					$link_cont=~s#^\s*$##gsi;
					if ($link_cont eq '')
					{
							if ($cont=~m#<body[^\>]*?>\s*<(\w+)[^\>]*?>(.*?)</\1>\s*<(\w+)[^\>]*?>(.*?)</\3>#si)
							{
								$link_cont=$4;
								$link_cont=~s#<[^\>]*?>##gsi;
							}
					}
			}
			$ncx_cont.="<navPoint id=\"\" playOrder=\"\"><navLabel><text>$link_cont</text></navLabel><content src=\"html/$htmlname\"/></navPoint>\n";
			
			#-- For New TOC Creation
			$toc_cont.="\n".'<p class="toc"><a href="'.$htmlname.'">'.$link_cont.'</a></p>';
			
			#-- For Nav Creation
			#<li><a href="../html/"></a></li>
			$nav_toc.="\n".'<li><a href="../html/'.$htmlname.'">'.$link_cont.'</a></li>';
			
	}
	
	$ncx_cont.="</navMap>\n</ncx>";

	
#-------------------------------- Cover File Creation -------------------------------#
	

	my $Cover_content_type=$1 if ($TitleMapCont=~m#Cover\=([^\n]*?)$#mi);

	if ($ComboCover eq '')
	{
			print "\n\tCover Creation\n";
			($ComboCover)=('Cover.html');
			if ($coverimagepath ne '')
			{
				copy("$coverimagepath","$IMGPath/cover.jpg");
			}

			my $NewCover_cont=$html_meta;
			$NewCover_cont=~s#<title>.*?</title>#<title>$Book_Title</title>#gsi;
			$NewCover_cont.="\n<body>";
			$NewCover_cont.="\n<div><img src=\"../img/cover.jpg\" alt=\"cover\"/></div>";
			$NewCover_cont.="\n</body>\n</html>";
			
			&WriteFile("$HtmlPath/$ComboCover", "$NewCover_cont", "Cover");
			
			$opf_cont=~s#(<item [^\>]*?media-type="application\/xhtml\+xml"[^\>]*?>)#\n<item id="cover-page" href="html/$ComboCover" media-type="application/xhtml\+xml"/>\n$1#is;
			if ($opf_cont=~m# id="cover-image"#si)
			{
				$opf_cont=~s#<item[^>]*?id\=\"cover\-image\"[^>]*?>#\n<item id="cover\-image" href="img/cover.jpg" media-type="image/jpeg"/>#is;
			}
			elsif($opf_cont=~m# media-type="image/jpeg"#si)
			{
				$opf_cont=~s#(<item [^\>]*?media-type="image/jpeg"[^\>]*?>)#\n<item id="cover\-image" href="img/cover.jpg" media-type="image/jpeg"/>\n$1#is;
			}
			else
			{
				$opf_cont=~s#(</manifest>)#\n<item id="cover\-image" href="img/cover.jpg" media-type="image/jpeg"/>\n$1#is;
			}
			$opf_cont=~s#(<spine toc="ncx">)#$1\n<itemref idref="cover-page"/>#is;
			$opf_cont=~s#<reference type="cover" title="Cover" href="html/"/>#<reference type="cover" title="Cover" href="html/$ComboCover"/>#is;
			
			
			if ($ncx_cont!~m#\/$ComboCover#si)
			{
				$ncx_cont=~s#<navMap>#<navMap>\n<navPoint id=\"nav1\" playOrder=\"1\"><navLabel><text>$Cover_content_type</text></navLabel><content src=\"html/$ComboCover\"/></navPoint>#igs;
			}
	}
=com
	else
	{
			#-- Change e.g. "1_1.jpg" to "cover.jpg" in html images and opf
			my $coverhtmlcont=&ReadFileDec("$HtmlPath/$ComboCover", "HTML");
			my $coverimagenew;
			if ($coverhtmlcont=~m/src\=\"\.\.\/img\/([^\"]*?)\"/is)
			{
				$coverimagenew=$1;
			}
			$coverhtmlcont=~s#(<img [^\>]*?img\/)[^\"]*?(\.\w+"[^\>]*?>)#${1}cover${2}#is;

			&WriteFile("$HtmlPath/$ComboCover", "$coverhtmlcont", "HTML");
			
			my $coverimageCopy=$coverimagenew;
			$coverimageCopy=~s#^.*?\.(\w+)$#cover\.$1#gsi;
			if ($coverimageCopy ne '')
			{
				rename("$IMGPath/$coverimagenew", "$IMGPath/$coverimageCopy");
				$opf_cont=~s/\b$coverimagenew/cover.jpg/igs;
			}
			
			#<text>[^\>]*?</text>\s*</navLabel>\s*<content src=\"[^\"]*?\/Title_Page.xhtml\"\s*/>\s*</navPoint>
	}
=cut
#-------------------------------- TOC File Creation -------------------------------#

	my $TOC_File="toc.html";
	my $toc_content_type=$1 if ($TitleMapCont=~m#contents\=([^\n]*?)$#mi);
	my $toc_fullCont;

	if ($TOCfiles[0] eq '')
	{
			print "\n\tTOC Creation\n";
			$toc_fullCont.=$html_meta;
			$toc_fullCont.="\n<body>";
			$toc_fullCont.="\n".'<p class="toc_title"><strong>'.$toc_content_type.'</strong></p>';
			$toc_fullCont.="$toc_cont";
			$toc_fullCont.="\n</body>\n</html>";
			$toc_fullCont=~s#<title></title>#<title>$Book_Title</title>#gsi;
			&WriteFile_DecToChar("$HtmlPath/$TOC_File", "$toc_fullCont", "TOC");
			
			#-- OPF - Insert After "begin chapter"
			if ($opf_cont=~m#<item id="([^\"]*?)"[^\>]*?href="html/$Begin_Chapter"[^>]*?>#si)
			{
				my $begin_opfId=$1;
				$opf_cont=~s#(<item [^\>]*?href="html/$Begin_Chapter"[^>]*?>)#<item id="file1" href="html/$TOC_File" media-type="application/xhtml\+xml"/>\n$1#is;
				$opf_cont=~s#(<itemref idref="$begin_opfId"\s*/>)#<itemref idref="file1"/>\n$1#is;
				$opf_cont=~s#(<reference type=\"toc\"[^>]*?) href=[^\>]*?>#$1 href=\"html/$TOC_File"/>#si;
			}
			
			#-- NCX - Insert After "begin chapter"
			$ncx_cont=~s#(<navPoint [^\>]*?>\s*<navLabel>\s*<text>((?:(?!<text|</text).)*?)</text>\s*</navLabel>\s*<content src="x?html/$Begin_Chapter"\s*/></navPoint>)#<navPoint id=\"nav1\" playOrder=\"1\"><navLabel><text>$toc_content_type</text></navLabel><content src=\"html/$TOC_File\"/></navPoint>\n$1#igs;
	}

#-------------------------------- NAV Creation -------------------------------#

print "\n\tNAV creation\n";

	my $cover_content_type=$1 if ($TitleMapCont=~m#cover\=([^\n]*?)$#mi);
	my $begin_content_type=$1;
	if ($Begin_Chapter ne '')
	{
			my $begin_cont=&ReadFileDec("$InPath/$Begin_Chapter", "BEGIN HTML");
			if ($begin_cont=~m#<body[^\>]*?>\s*<(\w+)[^\>]*?>(.*?)</\1>#si)
			{
					$begin_content_type=$2;
					$begin_content_type=~s#<[^\>]*?>##gsi;
					$begin_content_type=~s#^\s*$##gsi;
					if ($begin_content_type eq '')
					{
							if ($begin_cont=~m#<body[^\>]*?>\s*<(\w+)[^\>]*?>(.*?)</\1>\s*<(\w+)[^\>]*?>(.*?)</\3>#si)
							{
								$begin_content_type=$4;
								$begin_content_type=~s#<[^\>]*?>##gsi;
							}
					}
			}
	}

	#-- nav landmark
	my $nav_land;
	$nav_land.="<h2>Guide</h2>";
	$nav_land.="\n<ol>";
	$nav_land.="\n<li><a epub:type=\"cover\" href=\"../html/$ComboCover\">$cover_content_type</a></li>";
	if ($TOCfiles[0] ne '')
	{
		$nav_land.="\n<li><a epub:type=\"toc\" href=\"../html/$TOCfiles[0]\">$toc_content_type</a></li>";
	}
	else
	{
		$nav_land.="\n<li><a epub:type=\"toc\" href=\"../html/$TOC_File\">$toc_content_type</a></li>";
	}
	$nav_land.="\n<li><a epub:type=\"bodymatter\" href=\"../html/$Begin_Chapter\">$begin_content_type</a></li>";
	$nav_land.="\n</ol>";
	
	#-- nav content
	$nav_cont=~s#xml:lang=""#xml:lang="$Lang_code"#si;
	$nav_cont=~s#<title></title>#<title>$Book_Title</title>#si;
	$nav_cont=~s#<nav epub:type="toc"></nav>#<nav epub:type="toc">\n<ol>$nav_toc\n</ol>\n</nav>#si;
	$nav_cont=~s#<nav epub:type="landmarks"></nav>#<nav epub:type="landmarks">\n$nav_land\n</nav>#si;
	
#----------------------------- Copying Html Files ---------------------------#

	print "\n\tCopying HTMLs...\n";
	
	foreach my $htmlname (@htmlfiles)
	{
			my $cont=&ReadFileDec("$InPath/$htmlname", "HTML");
			$cont=~s#"epub.css"#"../css/epub.css"#gsi;
			$cont=~s#src="img/#src="../img/#gsi;
			&WriteFile_DecToChar("$HtmlPath/$htmlname", "$cont", "HTML");
	}
#----------------------------- Copying Images ---------------------------#
	
	if (-d "$InPath/img")
	{
			print "\n\tCopying Images...\n";
			dircopy("$InPath/img","$IMGPath") or die("$!\n");
	}
	
	#----------------------------- html to xhtml for Epub 3.0 ---------------------------#
	if ($Version eq '3.0')
	{
			print "\n\tEpub 3.0\n";
			
			opendir(DIR,$HtmlPath) or die("Unable to open html folder $HtmlPath: $!");
			my @newhtmlfiles=grep{/\.(html)$/i} readdir(DIR); closedir DIR;
			
			my @RenamedFile;
			foreach (@newhtmlfiles)
			{
					my ($old, $new)=($_, $_);
					$new=~s#\.html#\.xhtml#si;
					push (@RenamedFile,"$old<>$new");
			}
			
			foreach (@newhtmlfiles)
			{
					my ($htmfile, $newhtml)=($_, $_);
					$newhtml=~s#\.html#\.xhtml#gsi;
					my $NewHtmCont=&ReadFileDec("$HtmlPath/$htmfile", "HTML");
					foreach (@RenamedFile)
					{
							my $line=$_;
							my ($oldfile,$newfile)=($1,$2) if ($line=~m#^(.*?)<>(.*?)$#si);
							$NewHtmCont=~s#$oldfile#$newfile#gsi;
							$NewHtmCont=~s{\&\#65279;}{}gsi;
							$NewHtmCont=~s#<\!DOCTYPE [^\>]*?>\n?##gsi;
							$NewHtmCont=~s#<html [^\>]*?>#<html lang="$Lang_code" xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops">#gsi;
							&WriteFile_DecToChar("$HtmlPath/$newhtml", "$NewHtmCont", "HTML");
							
							$opf_cont=~s#$oldfile#$newfile#gsi;
							$ncx_cont=~s#$oldfile#$newfile#gsi;
							$nav_cont=~s#$oldfile#$newfile#gsi;
							
							if (-f "$HtmlPath/$newhtml")
							{
								unlink("$HtmlPath/$htmfile");
							}
					}
			}
			
			&WriteFile_DecToChar("$NAVPath/nav.xhtml", "$nav_cont", "NAV");
			
			my ($TYr,$TMon,$TDay,$THour,$TMin,$TSec) = getDateTime("4");
			$opf_cont=~s#<meta property="dcterms:modified">.*?</meta>#<meta property="dcterms:modified">${TYr}\-${TMon}\-${TDay}T${THour}\:${TMin}\:${TSec}Z</meta>#gsi;
	}
	
#----------------------------- Write OPF NCX ---------------------------#

	#-- OPF
	if ($Version eq '2.0')
	{
			$opf_cont=~s#<package xmlns:dc="http://purl.org/dc/elements/1.1/"#<package#gsi;
			$opf_cont=~s#version="3.0"#version="2.0"#gsi;
			$opf_cont=~s#<item [^\>]*?properties="nav"\s*[^\>]*?>\n?##si;
			$opf_cont=~s#<meta property="dcterms:modified">.*?</meta>\n?##gsi;
	}
	
	&WriteFile_DecToChar("$OEBPS/content.opf", "$opf_cont", "OPF");

	#-- NCX
	my $navcount='1';
	$ncx_cont=~s# id\=\"[^\"]*?\"#' id="nav'.$navcount++.'"'#iges;
	my $playcount='1';
	$ncx_cont=~s# playOrder\=\"[^\"]*?\"#' playOrder="'.$playcount++.'"'#iges;

	&WriteFile_DecToChar("$OEBPS/toc.ncx", "$ncx_cont", "NCX");

#----------------------------- Creating Epub ---------------------------#

		my $Source_Dir="$EpubFolder";
		my $Dest_Epub="$EpubFolder.epub";

		if (-f "$Dest_Epub")
		{
			rename("$Dest_Epub", "${EpubFolder}.Backup");
		}
		my $zip = Archive::Zip->new();

		# &zipFiles($Source_Dir,$Dest_Dir);
		my $string_member = $zip->addString( 'application/epub+zip', 'mimetype' );
		$string_member->desiredCompressionMethod( COMPRESSION_STORED );
		unlink("$Source_Dir/mimetype");
		$zip->addTree($Source_Dir,'');
		unless ( $zip->writeToFileNamed("${Dest_Epub}") == AZ_OK ) {
			   die Win32::MsgBox("",0,"S4C");
		}

		open (OUT, ">$Source_Dir/mimetype");
		print OUT 'application/epub+zip';
		close OUT;

#----------------------------- Sub Functions ---------------------------#
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
sub WriteFile_UTF8
{
	my $outfile=shift;
	my $cont=shift;
	my $type=shift;
	open (OUT,'>:utf8', "$outfile") or die("Unable to write $type file $outfile: $!");
	print OUT $cont;
	close OUT;
}
sub getDateTime
{
	my($YrValue) = @_;
	my ($sec,$min,$hour,$day,$mon,$year)=localtime(time);
	$day= sprintf ("%02d", $day);
	$hour= sprintf ("%02d", $hour);
	$min= sprintf ("%02d", $min);
	$sec= sprintf ("%02d", $sec);
	my $month=$mon+1;	$month= sprintf ("%02d", $month);
	$year=$year+1900;
	$year = substr $year, -2 if($YrValue eq "2");
	return ($year,$month,$day,$hour,$min,$sec);
}
