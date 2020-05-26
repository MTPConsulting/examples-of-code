<?php

namespace App\Services;

use App\Job;

class JobService
{
    /**
     * Get job
     *
     * @param $id
     * @param $slug
     * @return \Illuminate\Database\Eloquent\Collection|404
     */
    public function getJob($id, $slug)
    {
        return Job::where('id', $id)->where('slug', $slug)->with('subcategories')->firstOrFail();
    }

    /**
     * Get jobs related to category
     *
     * @param $id
     * @return \Illuminate\Contracts\Pagination\LengthAwarePaginator
     */
    public function getJobsCategory($id)
    {
        $jobs = Job::where('category_id', $id)
            ->orderBy('created_at', 'DESC')
            ->with('category')
            ->with('subcategories')
            ->with('user')
            ->paginate(15);

        return $jobs;
    }

    /**
     * Get jobs related to subcategory
     *
     * @param $id
     * @param $subcategory_id
     * @return \Illuminate\Contracts\Pagination\LengthAwarePaginator
     */
    public function getJobsSubCategory($id, $subcategory_id)
    {
        $jobs = Job::where('category_id', $id)
            ->whereHas('subcategories', function($q) use($subcategory_id) {
                $q->where('subcategory_id', $subcategory_id);
            })
            ->with('category')
            ->with('subcategories')
            ->with('user')
            ->orderBy('created_at', 'DESC')
            ->paginate(15);

        return $jobs;
    }
}