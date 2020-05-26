<?php

namespace Seny\Paginate;

use Seny\Paginate\Configuration;
use Seny\Paginate\Exception;
use Seny\Paginate\LoaderInterface;

/**
 * @author Martín Peveri <martinpeveri@gmail.com>
 */
final class Pager
{
    /**
     * @var LoaderInterface
     *
     * Loader for work
     */
    private $loader;

    /**
     * @var Configuration
     *
     * Configuration object
     */
    private $config;

    /**
     * @var int
     *
     * Current page
     */
    private $page = 1;

    /**
     * @var int
     *
     * Current position array
     */
    private $position = 0;

    public function __construct(LoaderInterface $loader, Configuration $config)
    {
        $this->loader = $loader;
        $this->config = $config;
    }

    public function getPageSize() : int
    {
        return $this->config->getSize();
    }

    public function getCurrentPage() : int
    {
        return $this->page;
    }

    public function getCurrentItems() : array
    {
        $size = $this->config->getSize();
        $page = $this->getCurrentPage();
        $data = $this->loader->getCurrentItems($size, $page, $this->position, $this->config->getItems());

        return $data;
    }

    public function getNextPage() : int
    {
        $page = $this->getCurrentPage();
        if ($page < $this->getTotalPages()) {
            $page++;
        }

        return $page;
    }

    public function getPreviousPage() : int
    {
        $page = $this->getCurrentPage();
        if ($page > 1) {
            $page--;
        }

        return $page;
    }

    public function next() : int
    {
        if ($this->page < $this->getTotalPages()) {
            $this->page++;
        }

        $this->setPage($this->page, true);

        return $this->page;
    }

    public function previous() : int
    {
        if ($this->page > 1) {
            $this->page--;
        }

        $this->setPage($this->page, true);

        return $this->page;
    }

    public function last() : int
    {
        $this->setPage($this->getTotalPages(), false);
        return $this->page;
    }

    public function first() : int
    {
        $this->setPage(1, false);
        return $this->page;
    }

    public function setPage(int $page, bool $restPage = false) : int
    {
        if ($page > 0 && $page <= $this->getTotalPages()) {
            $this->page = $page;

            $p = $page;
            if ($restPage) {
                $p -= 1;
            }

            $this->position = ($p * $this->config->getSize());
        } else {
            throw new Exception("Page $page does not exist.");
        }

        return $this->page;
    }

    /**
     * Get total pages
     *
     * @return int
     */
    public function getTotalPages() : int
    {
        return ceil($this->getTotalItems() / $this->config->getSize());
    }
    
    /**
     * Get total items
     *
     * @return int
     */
    public function getTotalItems() : int
    {
        return $this->loader->count($this->config->getItems());
    }

    public function getTotalItemsPage() : int
    {
        return $this->loader->count($this->getCurrentItems());
    }

    public function pageExists(int $page) : bool
    {
        $exists = false;
        if ($page > 0 && $page <= $this->getTotalPages()) {
            $exists = true;
        }

        return $exists;
    }
}
