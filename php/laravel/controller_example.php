<?php 

namespace App\Http\Controllers;

use DateTime;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\Input;
use App\Http\Requests;
use App\Category;
use App\Client;
use App\Participant;
use App\Participation;
use App\Country;

class RecruiterQueryController extends Controller 
{

    /**
    * Create a new controller instance.
    *
    * @return void
    */
    public function __construct()
    {
        $this->middleware('auth');
    }

    public function index()
    {   
        //Datos para los combos
        $combos = $this->getDataCombos();
        $categories = $combos['categories'];
        $clients = $combos['clients'];
        $countries = $combos['countries'];

        return view('recruiterquery.index')
            ->with("categories", $categories)
            ->with("clients", $clients)
            ->with("countries", $countries);
    }

    public function post(Request $request) 
    {
        //Valido el formulario
        $this->validateForm($request);
        //Obtengo el archivo
        $path = Input::file("xls_file")->getRealPath();

        //Leo los datos del archivo
        $excel = \App::make('excel');
        $data = $excel->load($path, function($reader) {
        })->get();

        $now = new DateTime();
        $records = [];
        $sheet1 = $data[0];

        foreach($sheet1 as $record) {
            //Obtengo el dni del participante
            $dni = $record['dni'];

            //Si es null no lo procesa
            if(is_null($dni)) {
                continue;
            }

            $participant = Participant::where('nro_documento', $dni)->first();
            if($participant) {
                //Si no esta habilitado o esta bloqueado no lo habilito
                if (!$participant->st_habilitado || $participant->st_blacklist) {
                    $habilitado = false;
                }else {
                    $habilitado = true;
                }

                //Si es menor a 6 mes con respecto a hoy
                if(isset($participant->ParticipationId->fecha)) {
                    //Obtengo los estudios cerrados del participante
                    $participations_user = Participation::where(
                        'xad_participantes_id', $participant->id
                    )->whereNotNull("fecha")->get();

                    foreach($participations_user as $p) { 
                        if($p->fecha != null) {
                            $input = new DateTime($p->fecha);
                            $diff = $input->diff($now);
                            //Si es menor a 6 mes con respecto a hoy
                            if($diff->y > 0 || $diff->m > 6) {
                                $habilitado = true;
                            } else {
                                $habilitado = false;
                            }
                        } else {
                            $habilitado = false;
                        }

                        if(!$habilitado) {
                            break;
                        }
                    }
                } else{
                    $habilitado = true;
                }

            } else {
                $habilitado = true;
            }

            $date_born = new DateTime($record->fecha_de_nacimiento->toDateTimeString());
            $interval = $date_born->diff($now);
            $edad = $interval->format('%y años');

            $record['habilitado'] = $habilitado;
            $record['edad'] = $edad;
            array_push($records, $record);
        }

        //Datos para los combos
        $combos = $this->getDataCombos();
        $categories = $combos['categories'];
        $clients = $combos['clients'];
        $countries = $combos['countries'];

        return view('recruiterquery.index')
            ->with("categories", $categories)
            ->with("clients", $clients)
            ->with("countries", $countries)
            ->with("records", $records);
    }

    public function UploadContentXls(Request $request) 
    {
        //Decodifico el json
        $data = json_decode($request->request->get('data'));
        $params = json_decode($request->request->get('params'));

        //Obtengo los parametros del form
        $client_id = $params->client_id;
        $category_id = $params->category_id;
        $number_study = $params->number_study;
        $country_id = $params->country_id;

        //Recorro los registros importados
        foreach($data as $record) {
            $nombre_apellido = $record->nombre_apellido;
            $dni = $record->dni;
            $fecha_nacimiento = $record->fecha_nacimiento;
            $nse = $record->nse;
            $habilitado = $record->habilitado;

            $participant_check = Participant::where('nro_documento', $dni)->first();
            if($participant_check != null) {
                $participant_check->nombre_apellido = $nombre_apellido;
                $participant_check->fec_nacimiento = $fecha_nacimiento;
                $participant_check->NSE = $nse;
                $participant_check->save();
                $participant = $participant_check;
            } else {
                $p = new Participant;
                $p->nombre_apellido = $nombre_apellido;
                $p->fec_nacimiento = $fecha_nacimiento;
                $p->NSE = $nse;
                $p->nro_documento = $dni;
                $p->st_habilitado = $habilitado;
                $p->st_blacklist = false;
                $p->user_id = Auth::user()->id;
                $p->save();
                $participant = $p;
            }
            
            //Si no está la participación la agrego
            if(Participation::where('estudio', $number_study)->where('nro_documento', $dni)->count() == 0) {
                $pr = new Participation;
                $pr->xad_participantes_id = $participant->id;
                $pr->nro_documento = $dni;
                $pr->estudio = $number_study;
                $pr->user_id = Auth::user()->id;
                $pr->xad_categorias_id = $category_id;
                $pr->xad_clientes_id = $client_id;
                $pr->id_pais = $country_id;
                $pr->save();
            }
        }

        return "Los postulantes han sido enviados exitosamente para su aprobación";
    }

    private function validateForm(Request $request) {
        $rules = [
            'number_study' => 'required',
            'xls_file' => 'required|mimes:xls,xlsx'
        ];  
        $niceNames = [
            'number_study' => 'Número de estudio',
            'xls_file' => 'Archivo xls'
        ]; 
        $this->validate($request, $rules, [], $niceNames);
    }

    private function getDataCombos() {
        $categories = Category::lists('detalle', 'id');
        $clients = Client::lists('razon_social', 'id');
        $countries = Country::lists('nm_pais', 'id_pais');
 
        return array(
            'categories' => $categories, 
            'clients' => $clients,
            'countries' => $countries
        );
    }

}
