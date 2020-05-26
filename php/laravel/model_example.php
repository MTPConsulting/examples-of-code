<?php

namespace App;

use Illuminate\Database\Eloquent\Model;

class Job extends Model
{
    public function category()
    {
        return $this->belongsTo('App\Category');
    }

    public function user()
    {
        return $this->belongsTo('App\User');
    }

    public function subcategories()
    {
        return $this->belongsToMany('App\SubCategory', 'jobs_subcategories', 'job_id', 'subcategory_id');
    }
}