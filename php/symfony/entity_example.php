<?php

namespace App\Entity;

use Doctrine\Common\Collections\ArrayCollection;
use Doctrine\Common\Collections\Collection;
use Doctrine\ORM\Mapping as ORM;

/**
 * @ORM\Entity(repositoryClass="App\Repository\PersonalRepository")
 */
class Personal
{
    /**
     * @ORM\Id()
     * @ORM\GeneratedValue()
     * @ORM\Column(type="integer")
     */
    private $id;

    /**
     * @ORM\Column(type="string", length=255)
     */
    private $surname;

    /**
     * @ORM\Column(type="string", length=255)
     */
    private $names;

    /**
     * @ORM\Column(type="string", length=255)
     */
    private $email;

    /**
     * @ORM\Column(type="string", length=50, nullable=true)
     */
    private $phone;

    /**
     * @ORM\Column(type="string", length=255)
     */
    private $file_number;

    /**
     * @ORM\ManyToOne(targetEntity="App\Entity\Company", inversedBy="personals")
     * @ORM\JoinColumn(nullable=false)
     */
    private $company;

    /**
     * @ORM\OneToMany(targetEntity="App\Entity\Register", mappedBy="personal", orphanRemoval=true)
     * @ORM\OrderBy({"datetime" = "DESC"})
     */
    private $registers;

    public function __construct()
    {
        $this->registers = new ArrayCollection();
    }

    public function getId(): ?int
    {
        return $this->id;
    }

    public function getSurname(): ?string
    {
        return $this->surname;
    }

    public function setSurname(string $surname): self
    {
        $this->surname = $surname;

        return $this;
    }

    public function getNames(): ?string
    {
        return $this->names;
    }

    public function setNames(string $names): self
    {
        $this->names = $names;

        return $this;
    }

    public function getEmail(): ?string
    {
        return $this->email;
    }

    public function setEmail(string $email): self
    {
        $this->email = $email;

        return $this;
    }

    public function getPhone(): ?string
    {
        return $this->phone;
    }

    public function setPhone(?string $phone): self
    {
        $this->phone = $phone;

        return $this;
    }

    public function getFileNumber(): ?string
    {
        return $this->file_number;
    }

    public function setFileNumber(string $file_number): self
    {
        $this->file_number = $file_number;

        return $this;
    }

    public function getCompany(): ?Company
    {
        return $this->company;
    }

    public function setCompany(?Company $company): self
    {
        $this->company = $company;

        return $this;
    }

    /**
     * @return Collection|Register[]
     */
    public function getRegisters(): Collection
    {
        return $this->registers;
    }

    public function addRegister(Register $register): self
    {
        if (!$this->registers->contains($register)) {
            $this->registers[] = $register;
            $register->setPersonal($this);
        }

        return $this;
    }

    public function removeRegister(Register $register): self
    {
        if ($this->registers->contains($register)) {
            $this->registers->removeElement($register);
            // set the owning side to null (unless already changed)
            if ($register->getPersonal() === $this) {
                $register->setPersonal(null);
            }
        }

        return $this;
    }
}
