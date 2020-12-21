# ========== listBgcolor.pl �̏o�̓t�H�[�}�b�g�ݒ� ==========

use strict;
use warnings;

sub format {
  my ($value, $color) = ($_[0], $_[1]);
  $value =~ s/\\/\\\\/g;
  $value =~ s/"/\\"/g;
  $value =~ s/\n/\\n/g;
  return sprintf("  bgColor[\"%s\"] = '#%s';\n", $value, $color);
}

1;