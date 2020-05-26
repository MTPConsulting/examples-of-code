<?php

namespace App\Controller\Api\V1;

use App\Service\AssistanceService;
use FOS\RestBundle\Controller\Annotations as Rest;
use FOS\RestBundle\Controller\FOSRestController;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpKernel\Exception\NotFoundHttpException;
use Symfony\Component\Routing\Annotation\Route;

class AssistanceController extends FOSRestController
{
    private $assistanceService;

    public function __construct(AssistanceService $assistanceService)
    {
        $this->assistanceService = $assistanceService;
    }

    /**
     * @Rest\Post("/mobile-login-check")
     */
    public function mobileLogin(Request $request)
    {
        try {
            $params = json_decode($request->getContent(), true);
            $email = $params['email'];
            $personal = $this->assistanceService->getPersonalId($email);
        } catch (\Exception $e) {
            return $this->json(['status' => 400, 'Message' => $e->getMessage()]);
        }

        return $this->json(['status' => 200, 'data' => $personal]);
    }

    /**
     * @Rest\Get("/get-movements-by-personal/{id}")
     */
    public function getMovementsByPersonal(Request $request, int $id)
    {
        try {
            $params = $request->query->all();
            $movements = $this->assistanceService->getMovementsByPersonal($id, $params);
        } catch (\Exception $e) {
            return $this->json(['status' => 400, 'Message' => $e->getMessage()]);
        }

        return $this->json(['status' => 200, 'data' => $movements]);
    }

    /**
     * @Rest\Post("/create-movement-type")
     */
    public function createMovementType(Request $request)
    {
        try {
            $params = json_decode($request->getContent(), true);
            $this->assistanceService->createMovementType($params);
        } catch (\Exception $e) {
            return $this->json(['status' => 400, 'Message' => $e->getMessage()]);
        }

        return $this->json(['status' => 201, 'Message' => 'Movimiento registrado con éxito']);
    }
}
