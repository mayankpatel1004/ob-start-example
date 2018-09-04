<?php
$border = "1";
$pagetitle = '<font size="+2">CONTACT REPORT</font>';
ob_end_clean();
ob_start();
?>
<div class="search_bar">
    <table width="702" border="0" height="50" cellspacing="0" cellpadding="0">
        <tr>
            <td>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td width="200">&nbsp;</td>
                        <td>
                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                    <td width="50%"><?php echo $pagetitle; ?></td>
                                </tr>
                            </table>
                        </td>
                        <td width="200">&nbsp;</td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</div>
<div class="clear"></div>
<div class="container01">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <td>
                <table width="100%" border="<?php echo $border; ?>" class="table_brd" cellspacing="1" cellpadding="0">
                    <tr bgcolor="#666666">
                        <th style="color:#FFFFFF;" valign="middle" width="10%" height="30">NAME</th>
                        <th style="color:#FFFFFF;" valign="middle" width="20%">EMAIL</th>
                        <th style="color:#FFFFFF;" valign="middle" width="10%">COMPANY</th>
                        <th style="color:#FFFFFF;" valign="middle" width="10%">PHONE</th>
                        <th style="color:#FFFFFF;" valign="middle" width="10%">SUBJECT</th>
                        <th style="color:#FFFFFF;" valign="middle" width="5%">TYPE</th>
                        <th style="color:#FFFFFF;" valign="middle" width="30%">MESSAGE</th>
                        <th style="color:#FFFFFF;" valign="middle" width="30%">STATUS</th>
                        <th style="color:#FFFFFF;" valign="middle" width="30%">SOURCE</th>
                    </tr>


                    <tr>
                        <td style="text-align:left;">Mayank</td>
                        <td style="text-align:left;">mayank.patel@devdigital.com</td>
                        <td style="text-align:left;">Dev Digital</td>
                        <td style="text-align:left;">999-999-9999</td>
                        <td style="text-align:left;">Test Subject</td>
                        <td style="text-align:left;">Test Type</td>
                        <td style="text-align:left;">Test Message</td>
                        <td style="text-align:left;">HOT</td>
                        <td style="text-align:left;">Source title</td>
                    </tr>

                </table>
            </td>
        </tr>
    </table>
    <br />
    <br />
</div>
<?php
$content = ob_get_contents();
$fp = fopen("contactreport.xls", "w");
fwrite($fp, $content, strlen($content));
fclose($fp);

header('Pragma: public');
header("Expires: Sat, 26 Jul 1997 05:00:00 GMT");
header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
header('Cache-Control: no-store, no-cache, must-revalidate');
header('Cache-Control: pre-check=0, post-check=0, max-age=0');
header("Pragma: no-cache");
header("Expires: 0");
header('Content-Transfer-Encoding: none');
header('Content-Type: application/vnd.ms-excel;');
header("Content-type: application/x-msexcel");
header("Content-Disposition: attachment;filename=contactreport.xls");
header("Content-Transfer-Encoding: binary ");
exit;
?>