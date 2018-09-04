package com.estimote.proximity;

import android.content.Context;
import android.content.Intent;
import android.content.SharedPreferences;
import android.os.Bundle;
import android.support.v7.app.AppCompatActivity;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.ProgressBar;
import android.widget.Toast;

import com.estimote.proximity.estimote.LoginTask;
import com.estimote.proximity.estimote.RegisterTask;
import com.estimote.proximity.estimote.Utils;

import org.json.JSONException;
import org.json.JSONObject;


public class MainActivity extends AppCompatActivity {

    private EditText surname, name, dni, email;
    private Button btnRegister;
    private ProgressBar progressBar;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        final SharedPreferences prefs = getSharedPreferences("auth", Context.MODE_PRIVATE);
        final String token = prefs.getString("CdPersona", "");
        if (!token.isEmpty()) {
            Intent intentHome = new Intent(getApplicationContext(), HomeActivity.class);
            startActivity(intentHome);
        }

        // Campos del formulario
        surname = findViewById(R.id.surname); // NmApellido
        name = findViewById(R.id.name); // NmNombres
        email = findViewById(R.id.email); // NmMail
        dni = findViewById(R.id.dni); // CdUsuario

        // ProgressBar
        progressBar = (ProgressBar) findViewById(R.id.progressBar);
        progressBar.setMax(10);
        progressBar.setVisibility(View.GONE);

        // Submit
        btnRegister = findViewById(R.id.button_submit);

        // Click Submit
        btnRegister.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                progressBar.setVisibility(View.VISIBLE);
                // Registro
                final String sUsername = surname.getText().toString().trim();
                final String sName = name.getText().toString().trim();
                final String sEmail = email.getText().toString().trim();
                final String sDni = dni.getText().toString().trim();

                Utils.context = getApplicationContext();
                final int duration = Toast.LENGTH_SHORT;

                if (sUsername.equals("") || sName.equals("") || sDni.equals("") || sEmail.equals("")) {
                    CharSequence text = "¡Debe completar los campos!";
                    Toast toast = Toast.makeText(Utils.context, text, duration);
                    toast.show();
                } else {
                    new LoginTask(new LoginTask.IEvents() {
                        @Override
                        public void AfterExecute(String response) throws JSONException {
                            JSONObject obj = new JSONObject(response);
                            if(obj.getBoolean("Ok")) {
                                String token_auth = obj.getJSONObject("Data").getString("token");
                                SharedPreferences.Editor editor = prefs.edit();
                                editor.putString("token", token_auth);
                                editor.apply();

                                if (token_auth != null && !token_auth.isEmpty()) {
                                    new RegisterTask(new RegisterTask.IEvents() {
                                        @Override
                                        public JSONObject BeforeExecute() throws JSONException {
                                            JSONObject jsonParam = new JSONObject();
                                            jsonParam.put("NmApellido", sUsername);
                                            jsonParam.put("NmNombres", sName);
                                            jsonParam.put("NmMail", sEmail);
                                            jsonParam.put("CdUsuario", sDni);
                                            jsonParam.put("CdCliente", Utils.cd_cliente);

                                            return jsonParam;
                                        }

                                        @Override
                                        public void AfterExecute(String response) throws JSONException {
                                            System.out.println(response);
                                            JSONObject obj = new JSONObject(response);
                                            if(obj.getBoolean("Ok")) {
                                                String cd_persona = obj.getJSONObject("Data").getString("CdPersona");
                                                System.out.println("Hola");

                                                System.out.println(cd_persona);
                                                SharedPreferences.Editor editor = prefs.edit();
                                                editor.putString("CdPersona", cd_persona);
                                                editor.apply();

                                                progressBar.setVisibility(View.GONE);
                                                Intent intentHome = new Intent(getApplicationContext(), HomeActivity.class);
                                                startActivity(intentHome);
                                            } else {
                                                progressBar.setVisibility(View.GONE);
                                                Toast toast = Toast.makeText(Utils.context, "Error al crear el usuario.", duration);
                                                toast.show();
                                            }
                                        }
                                    }).execute();
                                } else {
                                    progressBar.setVisibility(View.GONE);
                                    Toast toast = Toast.makeText(Utils.context, "Error al obtener el token.", duration);
                                    toast.show();
                                }
                            } else {
                                progressBar.setVisibility(View.GONE);
                                Toast toast = Toast.makeText(Utils.context, "Ocurrió un error al registrarse.", duration);
                                toast.show();
                            }
                        }
                    }).execute();

                }
            }
        });
    }
}
