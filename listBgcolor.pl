# ========== Excelファイルのセル色情報を抽出 ==========

use strict;
use warnings;
use utf8;
use Win32::LongPath;
use Encode;
use Win32::OLE;
use lib '.';
use listBgcolorFormat;

  my ($wd, $fn, $ex, $bk, %old);

  binmode STDOUT, ':crlf:encoding(cp932)';
  binmode STDERR, ':crlf:encoding(cp932)';
  binmode STDIN, ':crlf:encoding(cp932)';

  # カレントディレクトリを求める
  $wd =  getcwdL();
  $wd .= "\\" if (substr($wd, -1, 1) ne "\\");

  if (@ARGV == 1) {
    $fn = Encode::decode('cp932', $ARGV[0]);
  }
  else {
    print STDERR 'Usage: ' . Encode::decode('cp932', $0) . " ファイル名\n";
    exit 1;
  }

  unless ($ex = Win32::OLE ->
      new('Excel.Application', sub { $_[0] -> Quit; })) {
    print STDERR "Excelを起動できません\n";
    exit 1;
  }

  # 相対パスを絶対パスに変換
  if ($fn =~ /^[A-Za-z]:/) {
    ;
  }
  elsif ($fn =~ /^\\\\/) {
    ;
  }
  elsif ($fn =~ /^\\/) {
    print STDERR "ファイル名の指定が不適切です\n";
    exit 1;
  }
  else {
    $fn = $wd . $fn;
  }

  unless ($bk = $ex -> Workbooks -> Open(Encode::encode('cp932', $fn))) {
    print STDERR $fn . " を開けません\n";
    exit 1;
  }

  %old = ();

  for (my $i = 1; my $sht = $bk -> Worksheets($i); $i ++) {
    my ($ur, $ny, $nx);
    $ur = $sht -> UsedRange;
    ($nx, $ny) = ($ur -> Columns -> {'Count'}, $ur -> Rows -> {'Count'});

    for (my $y = 1; $y <= $ny; $y ++) {
      for (my $x = 1; $x <= $nx; $x ++) {
        my ($cell, $val, $in, $c, $key, $bg);
        $cell = $ur -> Cells($y, $x);
        $val = $cell -> {'Value'};
        next if (!defined($val) || $val eq '');
        $in = $cell -> Interior;
        next if ($in -> {'ColorIndex'} == -4142);  # 色なし
        $c = sprintf("%06x", $in -> {'Color'});

        $val = Encode::decode('cp932', $val);
        $key = $c . $val;
        next if (defined($old{$key}));
        $old{$key} = 1;

        $bg = substr($c, 4) . substr($c, 2, 2) . substr($c, 0, 2);

        print &format($val, $bg);
      }
    }

  }

