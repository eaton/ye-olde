#!/usr/bin/perl
#
# "Book 'em, Dan-O" Logger Script by Spider
# Created 9 July 1996
# Email me at  spider@servtech.com
# http://www.servtech.com/public/spider
#
# Secret Decoder Ring - aka Organization of Log File
# Time Stamp / Person\Machine / Referring URL / Browser Used
# 
# This script can be run as a SSI or used 
# in a "redirect" fashion via *normal* CGI calls.

########## Set Variables ############

$SSI = 1;  
# 0 if not used as a SSI  -    1 if used as a SSI

$logfile = "data/dan_o.dat";  
# change the directory path, silly!

$exclude = 0; 
# 1 if you want to exclude YOUR IP/Domain/Machine Name  - 0 otherwise

$my_addr = "your.machine.name";  
# used with the "exclude" portion

$HomeDirURL = "http://miso.wwa.com/~eaton/";  
# change this if you're not using SSI's

$nextfile = "index.html";  
# again, change if you're not using SSI's 

########## So much for that.. On with the show! #######

# Get the input
read(STDIN, $buffer, $ENV{'CONTENT_LENGTH'});

# Split the name-value pairs
@pairs = split(/&/, $buffer);

foreach $pair (@pairs) {
    ($name, $value) = split(/=/, $pair);

    # Un-Webify plus signs and %-encoding
    $value =~ tr/+/ /;
    $value =~ s/%([a-fA-F0-9][a-fA-F0-9])/pack("C", hex($1))/eg;

    # Stop people from using subshells to execute commands
    # Not a big deal when using sendmail, but very important
    # when using UCB mail (aka mailx).
    $value =~ s/~!/ ~!/g; 

    # Uncomment for debugging purpose
    # print "Setting $name to $value<P>";
    $FORM{$name} = $value;
}

($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime(time);

if ($sec < 10) {
	$sec = "0$sec";
}
if ($min < 10) {
	$min = "0$min";
}
if ($hour < 10) {
	$hour = "0$hour";
}
if ($mon < 10) {
	$mon = "0$mon";
}
if ($mday < 10) {
	$mday = "0$mday";
}

$month = ($mon + 1);
@months = ("January","February","March","April","May","June","July","August","September","October","November","December");
$date = "$hour\:$min\:$sec $month/$mday/$year";

# Now that we know what the time/date is.. let's have fun

if ($SSI == 1) {
	if ($exclude == 1) {
		&log unless ($ENV{'REMOTE_HOST'} eq $my_addr);
	} else {
	&log;
	}
	exit;
}

if ($SSI == 0) {
	if ($exclude == 1) {
		&log unless ($ENV{'REMOTE_HOST'} eq $my_addr);
	} else {
	&log;
	}
	&redir;
	exit;
}

sub log {

	if (! open(LOG,">>$logfile")) {
		print "Content-type: text/html\n\n";
		print "Couldn't open $logfile so I'm bugging out..\n";
		exit;
	}
	print LOG "At $date, $ENV{'REMOTE_HOST'} came here from $ENV{'HTTP_REFERER'} using $ENV{'HTTP_USER_AGENT'}.\n";
	close (LOG);
}

sub redir {
	print "Location: $HomeDirURL$nextfile\n\n";
}

