#!/usr/local/bin/perl

use Cwd;

$dir = getcwd;

if($^O eq "MSWin32"){
    $sep = "\\";
} else {
    $sep = "/";
}

print "Content-type: text/html\n\n";

#Get current user dir. Stored in $dir
$index = index($dir, user) + 5;
$rel_path = substr($dir, $index);
if($^O eq "MSWin32"){
	$dir =~ s/htdocs\/cgi-bin/htdocs/;
} else {
	$dir =~ s/cgi-bin/htdocs/;
}
$dir = $dir.$sep."awstats";

#Open $dir and read subdirectories, stored in @dirs
opendir(DIR, $dir) || die "can't opendir $dir: $!";
@dirs = grep { !/^\./ && -d "$dir/$_" } readdir(DIR);
closedir DIR;

if($rel_path =~ m/(\d+)/){
    $rel_path = $1;
}

($Second, $Minute, $Hour, $Day, $Month, $Year, $WeekDay, $DayOfYear, $IsDST) = localtime(time);
$Year -= 100;
$Month += 1;

if (length($Month) < 2){
  $Month = "0".$Month;
}

if (length($Year) < 2){
  $Year = "0".$Year;
}

$cur_date = $Year.$Month;


foreach (@dirs){

    if ( ($_ =~m/[0-9]/) && length($_) == 4)
    {
	my $label_year = convertYear($_);
	my $label_month = convertName($_);

    	$url = $_.$sep."awstats.".$rel_path.".".$_.".html";
    	$url =~ s/\\/\//g;

    	if($_ eq $cur_date){
        	print "<Option selected value=\"$url\">$label_month, $label_year\n";
    	} else {
        	print "<Option value=\"$url\">$label_month, $label_year\n";
    	}
    }
}

sub convertYear
{
	my $value = shift;

	$value = substr($value,0,2);
	return "20" . $value;
}

sub convertName
{
	my $value = shift;

	$value = substr($value,2,2);
	
	if ($value eq "01"){
		return "January";
	}
	elsif ($value eq "02"){
		return "February";
	}
	elsif ($value eq "03"){
		return "March";
	}
	elsif ($value eq "04"){
		return "April";
	}
	elsif ($value eq "05"){
		return "May";
	}
	elsif ($value eq "06"){
		return "June";
	}
	elsif ($value eq "07"){
		return "July";
	}
	elsif ($value eq "08"){
		return "August";
	}
	elsif ($value eq "09"){
		return "September";
	}
	elsif ($value eq "10"){
		return "October";
	}
	elsif ($value eq "11"){
		return "November";
	}
	elsif ($value eq "12"){
		return "December";
	}															
	
}
