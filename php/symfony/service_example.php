<?php

namespace App\Service;

use Doctrine\ORM\EntityManagerInterface;
use Symfony\Component\HttpKernel\Exception\NotFoundHttpException;
use Symfony\Component\Validator\Validator\ValidatorInterface;
use App\Entity\Register;
use App\Repository\PersonalRepository;
use App\Repository\RegisterRepository;

final class AssistanceService
{
    /**
     * @var PersonalRepository
     */
    private $personalRepository;

    /**
     * @var RegisterRepository
     */
    private $registerRepository;

    /**
     * @var EntityManagerInterface
     */
    private $em;

    /**
     * @var ValidatorInterface
     */
    private $validator;

    public function __construct(
        EntityManagerInterface $em,
        ValidatorInterface $validator,
        PersonalRepository $personalRepository,
        RegisterRepository $registerRepository
    ) {
        $this->em = $em;
        $this->validator = $validator;
        $this->personalRepository = $personalRepository;
        $this->registerRepository = $registerRepository;
    }

    /**
     * Obtiene el id de un personal, en base a un email
     *
     * @param string $email Email a buscar
     * @return array
     */
    public function getPersonalId(string $email)
    {
        $personal = $this->personalRepository->findBy(['email' => $email]);

        if (null === $personal) {
            throw new NotFoundHttpException('Not Found', null, 404);
        }

        return ["id" => $personal[0]->getId()];
    }

    /**
     * Obtiene los movimientos/registros de un personal
     *
     * @param int $id Identificador del personal
     * @param array $params Parametros datefrom y dateto
     * @return Array
     */
    public function getMovementsByPersonal($id, $params)
    {
        $registers = $this->registerRepository->filterByDatetime(
            $id,
            $params['dateFrom'],
            $params['dateTo']
        );

        $movements = [];
        foreach ($registers as $movement) {
            $movements[] = [
                "place" => $movement->getPlace(),
                "datetime" => $movement->getDatetime(),
                "movementType" => $movement->getMovementType()
            ];
        }

        return $movements;
    }

    /**
     * Crea un movimiento, de un personal, en la tabla de registros
     *
     * @return void
     */
    public function createMovementType(array $data)
    {
        $personal = $this->personalRepository->find($data['id']);

        $register = new Register();
        $register->setPersonal($personal);
        $register->setPlace($data['place']);
        $register->setDatetime(new \DateTime());
        $register->setMovementType($data['movementType']);

        $this->em->persist($register);
        $this->em->flush();
    }
}
