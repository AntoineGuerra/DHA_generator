<table style="max-width:368px;width:100%;" border="0" cellpadding="0" cellspacing="0" width="100%" class="copyable">
    <tbody>

    <?php if ($_GET['cdt']): ?>
    <tr>
        <td colspan="2">
            <span style="color: rgb(19,79,92); font-family: Helvetica, Arial,sans-serif; font-size: 13px;"><?php echo $_GET['cdt'] ?></span>
        </td>
    </tr>
    <tr>
        <td  colspan="2" style="font-family:Helvetica, Arial,sans-serif;font-size:15px;color:rgb(112,113,115);padding-bottom:20px;line-height:18px;border-bottom-style:solid;border-bottom-width:1px;border-bottom-color:rgb(112,113,115);" >
          <span style="display:inline-block;margin:20px 0 0 0;color:#383d46;font-size:17px;border-left:4px solid #18cdbe; padding:5px 15px;"><span style="font-weight:bold;font-size:18px;"><?php echo $_GET['name'] ?></span>
          <span style="display:block;color:#767d8b;font-size:15px;"> <?php echo $_GET['function'] ?></span>
        </span><br/>
            <?php else: ?>
    <tr>
        <td  colspan="2" style="font-family:Helvetica, Arial,sans-serif;font-size:15px;color:rgb(112,113,115);padding-bottom:10px;line-height:18px;border-bottom-style:solid;border-bottom-width:1px;border-bottom-color:rgb(112,113,115);" >
          <span style="display:inline-block;margin:20px 0 0 ;color:#383d46;font-size:17px;border-left:4px solid #18cdbe; padding:5px 15px;"><span style="font-weight:bold;font-size:18px;"><?php echo $_GET['name'] ?></span>
          <span style="display:block;color:#767d8b;font-size:15px;"> <?php echo $_GET['function'] ?></span>
        </span><br/>
            <?php endif; ?>
            <span style="line-height:10px;padding:0;margin:0;display:block;margin-top: 10px;">
          <a href="tel:<?php echo $_GET['tel_ctc'] ?>" style="color:rgb(112,113,115);text-decoration:none;" >
            <span style="font-weight:normal;color:#383d46;line-height:20px;font-weight: bold;"><?php echo $_GET['tel_print'] ?></span>
          </a>
          <?php if ($_GET['linkedin']): ?>
              <a href="<?php echo $_GET['linkedin'] ?>" style="display:block;color:rgb(112,113,115);text-decoration:underline;text-decoration-color: #9b9ea2" target="_blank">
            	<span style="text-decoration:none;">
          			<span style="font-weight:normal;color:#383d46;line-height:20px; font-weight: bold;font-size:13px;">
            			<?php echo $_GET['linkedin']; ?>
      	     		</span>
            	</span>
          	</a>
          <?php endif; ?>
        </span>
        </td>
    </tr>
    <tr height="55">
        <td style="padding:5px 0 0 0;width:190px;">
            <a href="http://www.agencemayflower.com" style="text-decoration:none;border:0;" target="_blank">
                <img src="http://www.agencemayflower.com/img_signature_mail/2018/mayflower.png" style="display:block;text-decoration:none;border:0;">
            </a>
        </td>
        <td class="social-picto" style="padding:5px 0 0 0; float:right;">
            <a href="https://www.instagram.com/agence_mayflower" style="text-decoration:none;border:0;" target="_blank">
                <img src="http://www.agencemayflower.com/img_signature_mail/2018/igIc.png" style="display:inline-block;text-decoration:none;border:0;" >
            </a>
            <a href="https://www.facebook.com/AgenceMayflower" style="text-decoration:none;border:0;" target="_blank">
                <img src="http://www.agencemayflower.com/img_signature_mail/2018/fbIc.png" style="display:inline-block;text-decoration:none;border:0;" >
            </a>
            <a href="https://www.linkedin.com/company/agence-mayflower" style="text-decoration:none;border:0;" target="_blank">
                <img src="http://www.agencemayflower.com/img_signature_mail/2018/lnIc.png" style="display:inline-block;text-decoration:none;border:0;" >
            </a>
        </td>
    </tr>
    <tr>
        <td colspan="2">
        <span style="margin:0;color:rgb(112,113,115);line-height:20px;font-size:12px;font-family:Helvetica, Arial,sans-serif;padding:0 0 20px 0;display:block;">
          <span class="phone-line" style="font-size:13px;">Lyon
          <a href="tel:0472698240" style="margin-left:6px;color:rgb(112,113,115);text-decoration:none;" >
            <span style="font-weight:normal;color:rgb(112,113,115);line-height:20px;">+33 (0)4 72 69 82 40</span>
          </a>
          </span>
          <span class="vertical-separator" style="margin-left: 8px; margin-right:8px;">|</span>
          <span class="phone-line" style="font-size:13px;"> Paris
          <a href="tel:0144709157" style="margin-left:8px;color:rgb(112,113,115);text-decoration:none;" >
            <span style="font-weight:normal;color:rgb(112,113,115);line-height:20px;">+33 (0)1 44 70 91 57</span>
          </a>
        </span>
          <br />
        </span>
        </td>
    </tr>
    <tr>
        <td colspan="2">
            <a href="http://www.agencemayflower.com" style="text-decoration:none;border:0;" target="_blank">
                <?php
                $ext = '';
                $imgDirPath = '/var/www/html/mayflower_2018/img_signature_mail/bandeau/';
                $imgUrl = '';
                foreach (scandir($imgDirPath) as $key => $file) {
                    if (preg_match('/bandeau_image\.(.*)/', $file, $matches)) {
                        $ext = $matches[1];
                        $imgUrl = 'http://www.agencemayflower.com/img_signature_mail/bandeau/bandeau_image.' . $ext;
//                    die(var_dump('<pre>', $matches, scandir('/var/www/html/mayflower_2018/img_signature_mail/bandeau/'), '</pre>'));
                        break;
                    } else {
                        $imgUrl = 'https://www.agencemayflower.com/img_signature_mail/2018/signature_mail_img.gif';
                    }
                }
                ?>
                <img class="img-gif" src="<?php echo $imgUrl ?>" />
            </a>
        </td>
    </tr>
    </tbody>
</table>
<br>
<button id="copy" style="width: 100%;">Copié</button>
<script src="public/js/copyable.js"></script>

