<?php 
 
// Load the database configuration file 
include_once 'dbConfig.php'; 
 
// Include PhpSpreadsheet library autoloader 
require_once 'vendor/autoload.php'; 
use PhpOffice\PhpSpreadsheet\Reader\Xlsx; 
 
if(isset($_POST['importSubmit'])){ 
     
    // Allowed mime types 
    $excelMimes = array('text/xls', 'text/xlsx', 'application/excel', 'application/vnd.msexcel', 'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); 
     
    // Validate whether selected file is a Excel file 
    if(!empty($_FILES['file']['name']) && in_array($_FILES['file']['type'], $excelMimes)){ 
         
        // If the file is uploaded 
        if(is_uploaded_file($_FILES['file']['tmp_name'])){ 
            $reader = new Xlsx(); 
            $spreadsheet = $reader->load($_FILES['file']['tmp_name']); 
            $worksheet = $spreadsheet->getActiveSheet();  
            $worksheet_arr = $worksheet->toArray(); 
 
            // Remove header row 
            unset($worksheet_arr[0]); 
 
            foreach($worksheet_arr as $row){ 
                $gender_id = $row[0]; 
                $patient_type_id = $row[1]; 
                $patient_origin_id = $row[2]; 
                $regime_id = $row[3]; 
                $health_center_id = $row[4]; 
                $sender_health_center_id = $row[5];
                $code = $row[6]; 
                $register_number = $row[7]; 
                $firstname = $row[8]; 
                $lastname = $row[9]; 
                $regime_date = $row[10]; 
                $backup_code = $row[11]; 
                $register_date_string = $row[12]; 
 
                // Check whether patient already exists in the database with the same email 
                $prevQuery = "SELECT id FROM patient WHERE health_center_id = '".$health_center_id."'"; 
                $prevResult = $db->query($prevQuery); 
                 
                if($prevResult->num_rows > 0){ 
                    // Update patient data in the database 
                    $db->query(
                        "UPDATE patient SET 
                        gender_id = '".$gender_id."', 
                        patient_type_id = '".$patient_type_id."', 
                        patient_origin_id = '".$patient_origin_id."', 
                        regime_id = '".$regime_id."', 
                        health_center_id = '".$health_center_id."', 
                        sender_health_center_id = '".$sender_health_center_id."', 
                        code = '".$code."', 
                        register_number = '".$register_number."', 
                        firstname = '".$firstname."', 
                        lastname = '".$lastname."', 
                        regime_date = '".$regime_date."', 
                        backup_code = '".$backup_code."', 
                        register_date_string = '".$register_date_string."', 
                        WHERE code = '".$code."'"
                    ); 
                }else{ 
                    // Insert patient data in the database 
                    $db->query(
                        "INSERT INTO patient (
                            gender_id, 
                            patient_type_id, 
                            patient_origin_id, 
                            regime_id, 
                            health_center_id, 
                            code, register_number
                            )
                     VALUES (
                        '".$gender_id."', 
                        '".$patient_type_id."', 
                        '".$patient_origin_id."', 
                        '".$regime_id."', 
                        '".$health_center_id."',
                        '".$code."',
                        '".$register_number."',
                        '".$register_number."',
                        '".$register_number."',
                        '".$register_number."',
                        '".$register_number."',
                        '".$register_number."',
                        '".$register_number."',
                        '".$register_number."'"
                    ); 
                } 
            } 
             
            $qstring = '?status=succ'; 
        }else{ 
            $qstring = '?status=err'; 
        } 
    }else{ 
        $qstring = '?status=invalid_file'; 
    } 
} 
 
// Redirect to the listing page 
header("Location: index.php".$qstring); 
 
?>