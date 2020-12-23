# ========== Excelファイルのセル色情報を抽出 ==========

use strict;
use warnings;
use Cwd;
use Win32::OLE;
use lib '.';
use listBgcolorFormat;

  my ($wd, $fn, $ex, $bk, %old);

  # カレントディレクトリを求める
  $wd =  Cwd::getcwd();
  $wd =~ s/\//\\/g;
  $wd .= "\\" if (substr($wd, -1, 1) ne "\\");

  if (@ARGV == 1) {
    $fn = $ARGV[0];
  }
  else {
    die 'Usage: ' . $0 . ' ファイル名';
  }

  $ex = Win32::OLE -> new('Excel.Application', sub { $_[0] -> Quit; })
    or die 'Excelを起動できません';

  # 相対パスを絶対パスに変換
  if ($fn =~ /^[A-Za-z]:/) {
    ;
  }
  elsif ($fn =~ /^\\\\/) {
    ;
  }
  elsif ($fn =~ /^\\/) {
    die 'ファイル名の指定が不適切です';
  }
  else {
    $fn = $wd . $fn;
  }

  $bk = $ex -> Workbooks -> Open($fn) or die $fn . ' を開けません';

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

        $key = $c . $val;
        next if (defined($old{$key}));
        $old{$key} = 1;

        $bg = substr($c, 4) . substr($c, 2, 2) . substr($c, 0, 2);

        print &format($val, $bg);
      }
    }

  }

