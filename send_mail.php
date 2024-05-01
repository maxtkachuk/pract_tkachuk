<?php
require 'PHPMailer/PHPMailer.php';
require 'PHPMailer/Exception.php';
require 'PHPMailer/SMTP.php';
require 'PHPExcel/PHPExcel.php';

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

$excelFile = 'database/list.xlsx'; /* шлях до твого excel */

$mail = new PHPMailer();

$mail->isSMTP();
$mail->Host       = 'smtp.office365.com';
$mail->SMTPAuth   = true;
$mail->Username   = 'maxtkachuk20@outlook.com'; // твоя пошта
$mail->Password   = 'maksmaks2020';        // Пароль від твоєї пошти
$mail->SMTPSecure = 'tls';
$mail->Port       = 587;
$mail->CharSet = 'UTF-8';
$mail->SMTPDebug = 2;  // Консоль

$objPHPExcel = PHPExcel_IOFactory::load($excelFile);
$sheet = $objPHPExcel->getActiveSheet();
$rows = $sheet->toArray();
foreach ($rows as $row) {
    if (!empty($row[0]) && !empty($row[1])) {
        $name = $row[1];
        $email = $row[0];
        if (filter_var($email, FILTER_VALIDATE_EMAIL)) {
            $mail->addAddress($email, $name);
            $mail->setFrom('maxtkachuk20@outlook.com', 'Maksym Tkachuk'); // твоя пошта
            $mail->addReplyTo('maxtkachuk20@outlook.com', 'Maksym Tkachuk'); // твоя пошта

            $subject = "Реєстраційні дані для участі у бета-тестуванні";  // Тема листа

            // Текст листа
            $message = "Шановний(а), $name!\n\n";
            $message .= "Ми вдячні Вам, що ви зареєструвались для участі у бета-тестуванні нової Системи дистанційного навчання. Надсилаємо Вам реєстраційні дані.\n\n";
            $message .= "Ваш логін – 'вказати логін', пароль – 'вказати пароль' та адреса для входу – http://194.44.152.161/.\n\n";
            $message .= "За цими реєстраційними даними Ви можете уже зараз зайти та ознайомитись із Системою. Починаючи з наступного тижня (з 15 березня 2021 р.) ми організуємо ряд навчальних он-лайн семінарів щодо роботи у новій Системі та формату співпраці у межах тестування, про що повідомимо Вам окремо.\n\n";
            $message .= "У разі виникнення питань просимо звертатись за адресою dist@pnu.edu.ua\n\n";
            $message .= "З повагою\nЦентр дистанційного навчання та моніторингу освітньої діяльності";

            $mail->isHTML(false);
            $mail->Subject = $subject;
            $mail->Body    = $message;

            if (!$mail->send()) {
                echo "Помилка: " . $mail->ErrorInfo;
            }

            $mail->clearAddresses();

            header("location: success.php");
            exit();
            
        } else {
            echo "Помилка: Неприпустима адреса електронної пошти: $email";
        }
    }
}

?>
