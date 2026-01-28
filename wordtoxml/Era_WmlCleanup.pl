use strict;
#use warnings;
#use Win32::OLE;
use File::Basename;
#use constant wdOpenFormatAuto => 0;
use List::Util qw(reduce);

my ($Input,$Output,$Value,@Values,$Tmp);
$Input=$ARGV[0];

my @suffixes=(".docx",".doc");
my $FileName= basename($ARGV[1], @suffixes);

open(XML, $Input) || die "Can't open the $Input file $!";
{  local $/; $_=<XML>; $Tmp=$_;  }

#$Tmp =~ s{(.*?)<w:body>(.*?)</w:body>(.*?)}{<converter><w:body>$2</w:body></converter>}si;

	# $Tmp =~ s{<w:body>(.*?)</w:body>}{"<w:body>" . &BodyText($1) . "</w:body>"}esi;
	
	$Tmp=~s{<w:smartTag([^\>]*)\>}{}gsi;

	$Tmp=~s{<\/w:smartTag([^\>]*)\>}{}gsi;

	while($Tmp=~s#<w:rStyle w:val="([0-9]+[^"]*)"#<w:rStyle w:val="A$1"#gsi){}

	my $count ='1';

	$Tmp=~s{<w:rStyle w:val=\"CommentReference\"/>}{"<w:rStyle w:val=\"CommentReference".$count++."\"\/>"}segi;

	

	$Tmp=~s{<w:del w:id=\"([^\"]+)\" w:author=\"([^\"]+)\" w:date=\"([^\"]+)\"\/>}{}gsi;

	$Tmp=~s{<w:del w:id=\"([^\"]+)\" w:author=\"([^\"]+)\"\/>}{}gsi;

	$Tmp=~s{<w:del ([^\>]+)\>(.*?)<\/w:del>}{"<w:del $1>".&Delete($2)."<\/w:del>"}gesi;

	

		#$Tmp=~s{<w:instrText([^\>]*)\>((\s?)SET File\_Name\:([^\:]+)([^\<]*))\<\/w:instrText>}{<\/w:r><w:r><w:rPr><w:rStyle w:val="File_path"/>$4</w:rPr>}gsi;

		
		
		#$Tmp=~s{<w:instrText([^\>]*)\>((\s?)SET History\_Date\:([^\:]+)([^\<]*))\<\/w:instrText>}{<\/w:r><w:r><w:rPr><w:rStyle w:val="History_Date"/>$4</w:rPr>}gsi;

		


	sub Delete
	{
		my $Delete=shift;

		

		$Delete=~s{<w:delText([^\>]*)\>}{\[\[\[w:delText\]\]\]}gsi;

		$Delete=~s{<\/w:delText>}{\[\[\[\/w:delText\]\]\]}gsi;

		$Delete=~s{<[^\>]+\>}{}gsi;

		$Delete=~s{\[\[\[w:delText\]\]\]}{<w:delText>}gsi;

		$Delete=~s{\[\[\[\/w:delText\]\]\]}{<\/w:delText>}gsi;

		
		return $Delete;
	}


	$Tmp=~s{<w:instrText([^\>]*)\>((\s?)SET Table([^\<]*))\<\/w:instrText>}{<\/w:r><w:r><w:rPr><w:rStyle w:val="Tablecount"/>$2</w:rPr>}gsi;

	#$Tmp=~s{<w:instrText([^\>]*)\>((\s?)SET Smallcaps([^\<]*))\<\/w:instrText>}{<\/w:r><w:r><w:rPr><w:rStyle w:val="sc"/></w:rPr>}gsi;

	
	
	
	$Tmp=~s{(<w:r><w:rPr><w:rStyle w:val="Tablecount"/>([^\<]*)\<\/w:rPr><\/w:r>)(.*?)<w:tbl>}{$3<w:tbl>$1}gsi;

	
	#$Tmp=~s{<w:instrText([^\>]*)\>((\s?)SET \"([^\<]*))\<\/w:instrText>}{<\/w:r><w:r><w:rPr><w:rStyle w:val="citationref"/>$2</w:rPr>}gsi;


	# print $Tmp."\n";
	
	#$Tmp=~s{<w:instrText([^\>]*)\>(.*?)</w:instrText>}{}gsi;

	
	#$Tmp=~s{(<w:tbl>(.*?)<\/w:tbl>)}{&Table($1)}gesi;

	
	
	#$Tmp=~s{(<w:tbl>(.*?)<\/w:tbl>)}{&Table_Span($1)}gesi;

	sub Table
	{
		my $Table=shift;

		
		
		if($Table=~m{<w:r><w:rPr><w:rStyle w:val="Tablecount"/>(\s?)SET Table:(\d+)([^\<]+)\<\/w:rPr>}gsi)
		{
			my $Column_count=$2;

		
			
			$Table=~s{(<w:tr ([^\>]*)\>(.*?)<\/w:tr>)}{&Row($1,$Column_count)}gesi;

			
			my @Grid;
		if($Table=~m{<Column\-width=([^\>]+)\>}gsi)
		{
			my $Width=$1;
			
			my @W=split(":",$Width);
			foreach my $Value(@W)
			{
				
				$Value="<w:gridCol w:w=\"$Value\"/>";
				push(@Grid,$Value);
			}
			my $Final_Grid=join("",@Grid);
		$Table=~s{<w:gridCol w:w=\"([^\"]+)\"\/>}{}gsi;
		$Table=~s{<w:tblGrid>}{<w:tblGrid>$Final_Grid}si;
		$Table=~s{<Column\-width=([^\>]*)\>}{}gsi;
		}
		}
	my $Count='1';
	$Table=~s{<w:gridcol w:w=\"([^\"]+)\"\/>}{'<w:gridcol'.$Count++.' w:w="'.$1."\">"}egsi;

	if($Table=~m{<w:r><w:rPr><w:rStyle w:val="Tablecount"/>(\s?)SET Table:(\d+)([^\<]+)\<\/w:rPr>}gsi)
	{
			my $Column_count=$2;

			
			
			$Table=~s{(<w:tr ([^\>]*)\>(.*?)<\/w:tr>)}{&Row_Col($1,$Column_count)}gesi;


			while($Table=~m{<w:gridcol(\d+) w:w=\"(\d+)\">}gsi)
			{
				my $Col=$1;
				my $Width=$2;
				
				$Table=~s{<w:tc$Col(\/?)>}{<w:tc$Col width:=\"$Width\"$1>}gsi;
			}


			
			#$Table=~s{(<w:tc(\d+) width:=\"(\d+)\">(.*?)<\/w:tc>)}{&Width($1)}gesi;
			
	}
		
	#}

		
		$Table=~s{<w:gridcol(\d+) w:w=\"(\d+)\">}{<w:gridCol w:w=\"$2\"\/>}gsi;
		$Table=~s{<w:tc(\d+) width:=\"(\d+)\"\/>}{}gsi;
		$Table=~s{<w:tc(\d+) width:=\"(\d+)\">}{<w:tc>}gsi;
		$Table=~s{<w:tc(\d+)>}{<w:tc>}gsi;
		$Table=~s{\n}{}gsi;
		
		return $Table;
		
	}

sub Width
{
	my $Col=shift;
	my $Column=$Col;
	my $Width;

	if($Column=~m{<w:tcW w:w=\"(\d+)\"}gsi)
	{
		$Width=$1;
	}

	
	
	my $Sum;
	while($Col=~m{<w:tc(\d+) width:=\"(\d+)\"}gsi)
	{
		
		$Sum=$Sum+$2;
		
	}

	#my $Final_Value=join(',',@Sum);
	#print "Check".$Final_Value."\n";
	#my @e=split(",",$Final_Value);
	#my $Add = eval join '+', @e;
	my $Add=$Sum;

	#print $Width."\n";

	
	
	
	if($Add ne $Width)
	{
		
		
		$Col=~s{<w\:gridSpan w\:val=\"([^\"]+)\"\/>}{}si;
	}
	
	$Col=~s{<w:tc(\d+) width:=\"(\d+)\"\/>}{}gsi;
	$Col=~s{<w:tc(\d+) width:=\"(\d+)\">}{<w:tc>}gsi;
	$Col=~s{<w:tc(\d+)>}{<w:tc>}gsi;
	return $Col;
}
	
sub Row
{
	my $Row=shift;

	my $Col_Count=shift;
	my $Count=$Row=~s{<w:tc>}{<w:tc>}gsi;
my @Grid;
	if($Col_Count eq $Count)
	{
		
		$Row=~s{<w\:gridSpan w\:val=\"([^\"]+)\"\/>}{}gsi;

		my $Column_width;
		
		while($Row=~m{<w:tcW w:w=\"(\d+)\"}gsi)
		{
			my $Col=$1;
			
			
			push(@Grid,$Col);
		}
		
	}
	my $Final_Col=join(":",@Grid);
	
	return "<Column\-width\=".$Final_Col.">".$Row;
}

sub Table_Span
{
	my $Table=shift;

	
	
	
	return $Table;
}

sub Row_Col
{
	my $Row=shift;

	
	my $Col_Count=shift;
	my $Count=$Row=~s{<w:tc>}{<w:tc>}gsi;
my @Grid;
	if($Col_Count ne $Count)
	{
		my $Count='1';
		#$Row=~s{<w:tc>}{'<w:tc'.$Count++.'>'}egsi;
		
	}
	$Row=~s{(<w:tc>(.*?)<\/w:tc>)}{&Column($1)}gesi;

	my $Cou='1';
	$Row=~s{<w:tc(\/?)>}{'<w:tc'.$Cou++.$1.'>'}egsi;
	return $Row;
}

sub Column

{
	my $Col=shift;
	#my $count=shift;

	my @Grids;
	if($Col=~m{<w\:gridSpan w\:val=\"([^\"]+)\"\/>}gsi)
	{
		my $Value=$1;
		
		#$Value=$Value+$count;
		
		for(my $i=1;$i<$Value;$i++)
		{
			my $Grid="<w:tc\/>";
			push(@Grids,$Grid);		
		}

		
	}
	my $Final_Grid=join("\n",@Grids);
	

	$Col=~s{(<w\:gridSpan w\:val=\"([^\"]+)\"\/>)}{$1\n$Final_Grid}si;
	
	return $Col;
}

sub BodyText {
my $Tmp1 = $1;
       # $Tmp1 =~ s{w:val="([^\"]+)"}{"w:val=\"" . &StyleReplace($1) . "\""}gesi;
        $Tmp1 =~ s{w:val="([^\"]+)Char"}{w:val="$1"}gsi;
	$Tmp1 =~ s{w:val="Aa"}{w:val="A"}gsi;

return $Tmp1;
}



sub MergeStyles {
    my $LeftVal = $2;
    my $RightVal = $6;
    my $WholeData = $4;
if ($LeftVal eq $RightVal){
    $WholeData = "";
}

return $WholeData;
}


#        $Tmp =~ s{w:val="Aa"}{w:val="A"}gsi;
#        #$Tmp =~ s{<w:name w:val="§}{<w:name w:val="}gsi;
#        #$Tmp1 =~ s{w:styleId="([^\"]+)"}{"w:val=\"" . &StyleReplace($1) . "\""}gesi;
#        $Tmp =~ s{<w:sectPr (.*?)</w:sectPr>|<w:instrText>(.*?)</w:instrText>}{}gsi;
#       

sub StyleReplace {
    my $tmpRepl = $1;
    if ($tmpRepl =~m{^(index[0-9]|NG-NL[0-9]|NG-BL[0-9]|NG-UL[0-9]|NL[0-9]|BL[0-9]|UL[0-9]|OL[0-9]|TOC[0-9]|TableList[0-9]|TableGrid[0-9]|TableColumns[0-9]|TableColorful[0-9]|T[0-9]|ListNumber[0-9]|ListContinue[0-9]|ListBullet[0-9]|List[0-9]|Heading[0-9]|TableClassic[0-9]|sub[0-9]|ng-toc[0-9]|ng-regL[0-9]|ng-m[0-9]|ng-lawL[0-9]|ng-index[0-9]|ng-h-[0-9]|ng-h[0-9]|ng-fm-h[0-9]|ng-compL[0-9])}i) {
        $tmpRepl =~s{^(index|NL|BL|UL|OL|TOC|TableList|TableGrid|TableColumns|TableColorful|T|ListNumber|ListContinue|ListBullet|List|Heading|TableClassic|sub|ng-toc|ng-regL|ng-m|ng-lawL|ng-index|ng-h-|ng-h|ng-fm-h|ng-compL|Footnote)(\d)(\d+)}{$1$2}gsi;
        $tmpRepl =~s{^(index|NL|BL|UL|OL|TOC|TableList|TableGrid|TableColumns|TableColorful|T|ListNumber|ListContinue|ListBullet|List|Heading|TableClassic|sub|ng-toc|ng-regL|ng-m|ng-lawL|ng-index|ng-h-|ng-h|ng-fm-h|ng-compL|Footnote)0}{$1}gsi;
    }
    else {
        $tmpRepl =~s{^([^\d]+)\d+$}{$1}gsi;
    }
    return $tmpRepl;
}



close(XML);
    unlink($Input);
sleep(2);
    open(OUT, ">$Input") || die "Can't open the $Input file $!";
    print OUT $Tmp;
    close(OUT);


    
