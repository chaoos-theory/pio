# DISPLAY INFO
function display( [string]$subject, [string]$color , [int]$length)  {

	# REQUIRED LENGTH OF STRING
	$len = $length

	# STRINGS THAT ARE LONGER WILL BE CUT DOWN,
	# STRINGS THAT ARE TO SHORT WILL BE MADE LONGER
	if ( $subject.length -lt $length ){
		$toadd=$length-$subject.length;
		for ( $i=0; $i -lt $toadd; $i++ ){
			$subject=$subject+" ";
		}
		$len = $subject.length
	}
	else { $len = $length }
	
	$index=$index+1
	Write-host -ForegroundColor $color ((($subject).ToString()).Substring(0,$len)).ToUpper()
}

function size_adjust( [string]$text, [int]$limit, $left=$False) {

	if ( $text.length -lt $limit ){
		$toadd=$limit-$text.length;
		for ( $i=0; $i -lt $toadd; $i++ ){
			if ($left -eq $False){
				$text=$text+" ";
			}
			else {
				$text=" "+$text;
			}
		}
	}
	$result = ((($text).ToString()).Substring(0,$limit)).ToUpper()
	return $result
}
clear-host
while ($True) {


	$outlook = New-Object -comobject outlook.application
	$namespace = $outlook.GetNameSpace("MAPI")
	$DefaultFolder = $namespace.GetDefaultFolder(6)
	$Emails = $DefaultFolder.Items

	$h = get-host
	$win = $h.ui.rawui.windowsize
	$width = $win.width
	$height = $win.height

	$unread_counter = 0
	$counter = 0 



	$tick = '-'

	# Requirement satisfaction
	if ($width -lt 80){
		Write-host "Terminal width of at least 80 columns is required. Quiting."
		exit
	}
	if ($height -lt 6){
		Write-host "Terminal height of at least 6 rows is required. Quiting."
		exit
	}


	$sendername_len = 30
	$sent_len = 19
	$size_len = 8
	$additional = 15
	$subject_len = $width - ( $sendername_len + $sent_len + $size_len + $additional)

	$head_subject = size_adjust "SUBJECT" $subject_len
	$head_sender = size_adjust "FROM" $sendername_len
	$head_sent =  size_adjust "SENT ON" $sent_len
	$head_size = size_adjust "SIZE" $size_len $True

	Write-host "  |" $head_sender "|" $head_sent "|" $head_subject "|" $head_size

	Foreach ($Email in $Emails) {
		$counter = $counter + 1

		$subject = size_adjust ($Email.Subject) $subject_len
		$sendername = size_adjust ($Email.SenderName) $sendername_len
		$unread = $Email.UnRead
		$to = ' '
		$size = [math]::Round($Email.Size / 1000)
		$size_txt = size_adjust ($size.ToString() + " KB") $size_len $True

		if ($Email.To -ne $null -and ($Email.To).ToUpper() -like "*SMT*") {
			$to = '>'
		}

		$sent = size_adjust ($Email.SentOn) $sent_len

		#display ( $subject + $sendername + $sent + "`n" ) ("DarkGray") ($width)

		if ($counter -gt ($Emails.Count - ($height-2))) {

			if ($unread -eq $False) {
				Write-host -ForegroundColor 'DarkGray' $to "|" $sendername "|" $sent "|" $subject "|" $size_txt
			}
			else {
				Write-host -ForegroundColor 'White' $to "|" $sendername "|" $sent "|" $subject "|" $size_txt
			}

		}
		else {
			if ($tick -eq '-'){
				$tick = '\'
			}
			elseif ($tick -eq '\'){
				$tick = '|'
			}
			elseif ($tick -eq '|'){
				$tick = '/'
			}
			elseif ($tick -eq '/'){
				$tick = '-'
			}
			Write-host -nonewline "`r " $tick
		}

		if ($counter -eq ($Emails.Count - ($height-2)) ) {
			Write-host -nonewline "`r											`r"
		}

	}

	Start-Sleep -s 30
}
