<?php

$nombreAux = isset($_POST['nombre']);
$emailAux = isset($_POST['email']);
$telefonoAux = isset($_POST['telefono']);
$consultaAux = isset($_POST['consulta']);

if($nombreAux && $emailAux && $telefonoAux && $consultaAux){

	$emptyNombre = empty($_POST['nombre']);
	$emptyEmail = empty($_POST['email']);
	$emptyTelefono = empty($_POST['telefono']);
	$emptyConsulta = empty($_POST['consulta']);

	if(!$emptyNombre && !$emptyEmail && !$emptyTelefono && !$emptyConsulta){
		//Datos de destino
		$email_destino = "";

		//Obtengo los datos a enviar
		$nombre = $_POST['nombre'];
		$email = $_POST['email'];
		$telefono = $_POST['telefono'];
		$consultaAux = $_POST['consulta'];

		//Armo el email
		$consulta = "";
		$consulta = $consulta . "Nombre: " . $nombre . " <br>";
		$consulta = $consulta . "Email: " . $email . " <br>";
		$consulta = $consulta . "Tel: " . $telefono . " <br>";
		$consulta = $consulta . "Mensaje: " . $consultaAux . " <br>";

		$from = $email;
		$to = $email_destino;
		$name = "Example";
		$subject = "Email desde Example ONLine";
		$message = $consulta;

		//Envío el email
		$from_user = "=?UTF-8?B?".base64_encode($from)."?=";
	    $subject = "=?UTF-8?B?".base64_encode($subject)."?=";

	    $headers = "From: $from <$from>\r\n". 
	               "MIME-Version: 1.0" . "\r\n" . 
	               "Content-type: text/html; charset=UTF-8" . "\r\n"; 

	    $resultado = mail($to, $subject, $message, $headers); 
		if($resultado){
			echo '<h4 class="alert-heading"><i class="fa fa-info-circle"></i>' . " Se envió el email con éxito.</h4>";
		}else{
			echo '<h4 class="alert-heading"><i class="fa fa-info-circle"></i>' . " Error al enviar el email.</h4>";
		}
	}else{
		echo '<h4 class="alert-heading"><i class="fa fa-info-circle"></i>' . " Todos los campos del formulario son obligatorios.</h4>";
	}
}else{
	echo '<h4 class="alert-heading"><i class="fa fa-info-circle"></i>' . " Datos incorrectos.</h4>";
}


?>
