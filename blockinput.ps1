function block-tskeyboardinput {
$signature = @"
  [dllimport("user32.dll")]
  public static extern bool blockinput(bool fblockit);
"@

$block = add-type -memberdefinition $signature -name disableinput -namespace disableinput -passthru
$block::blockinput($true)
}

function unblock-tskeyboardinput {
$signature = @"
  [dllimport("user32.dll")]
  public static extern bool blockinput(bool fblockit);
"@

$unblock = add-type -memberdefinition $signature -name enableinput -namespace enableinput -passthru
$unblock::blockinput($false)
}

block-tskeyboardinput
start-sleep -seconds 10
unblock-tskeyboardinput