#!/usr/bin/perl
#
#

use strict;

my ($sec, $min, $hour, $day, $mon, $year) = localtime(time);
$year = 1900+$year;
$mon = 1+$mon;
my $dir = "dividend_$year-$mon-$day"."_"."$hour-$min-$sec";
mkdir $dir;

# Table of all URL
###################
 # today
my $table_url = "http://mops.twse.com.tw/mops/web/t05sr01_1";
 # yesterday
#$year = $year - 1911;
#$day = $day - 1;
#my $table_url = "http://mops.twse.com.tw/mops/web/t05st02?step=0&newstuff=1&firstin=1&year=$year&month=$mon&day=$day";


# company info URL
# http://mops.twse.com.tw/mops/web/ajax_t05sr01_1?step=1&COMPANY_ID=1340&SPOKE_DATE=20160324&SPOKE_TIME=225941&SEQ_NO=2
##################
 # today
my $company_info_url = "http://mops.twse.com.tw/mops/web/ajax_t05sr01_1?step=1&";
 # yesterday
#my $company_info_url = "http://mops.twse.com.tw/mops/web/ajax_t05st02?step=1&";

# step1: Get table of all
`wget "$table_url" -O $dir/table.html`;

# step2: Retrive company one by one
my $get_company = 0;
my $get_dividend = 0;
my $company_name = "unknown";
open IN, "<$dir/table.html" or die "cannot open $dir/table.html\n";
#open IN, "<table.html" or die "cannot open table.html\n";
while (<IN>) {
	if (/<td>(\d\d\d\d)<\/td>/) {
		$get_company = 1;
	}
	elsif ($get_company == 1) {
		if (/<td.+>(.+)<\/td>/) {
			$company_name = $1;
		}
		$get_company = 2;
	}
	elsif (/<td.+>(.+股利.+)<\/td>/) {
		my $info = $1;
		$get_dividend = 1;
		if ($info =~ /不分派/) {
			$get_dividend = 0;
			$get_company = 0;
		}
		elsif ($1 =~ /不發/) {
			$get_dividend = 0;
			$get_company = 0;
		}
		print "$company_name\t$info\n";
	}
	elsif (/.+SEQ_NO.value='(\d+)'.*SPOKE_TIME.value='(\d+)'.*SPOKE_DATE.value='(\d+)'.*COMPANY_NAME.value='(.+)'.*COMPANY_ID.value='(\d+)'.+/) {
		my $seq_no = $1;
		my $spoke_time = $2;
		my $spoke_date = $3;
		my $_company_name = $4;
		my $company_id = $5;
		#print "$seq_no,$spoke_time,$spoke_date,$company_id\n";
		my $url = $company_info_url."COMPANY_ID=".$company_id."&SPOKE_DATE=".$spoke_date."&SPOKE_TIME=".$spoke_time."&SEQ_NO=".$seq_no;
		my $out_file = "$company_id-$company_name-$spoke_date-$spoke_time-$seq_no.html";
		if ($get_company && $get_dividend) {
			print "$out_file\n";
			print "$url\n";
			sleep 3;
			`wget \'$url\' -O $dir/$out_file`;
			print "==========================\n"
		}
		$get_company = 0;
		$get_dividend = 0;
		my $company_name = "unknown";
	}
}
close IN;

# step3: Get company infor one by one

