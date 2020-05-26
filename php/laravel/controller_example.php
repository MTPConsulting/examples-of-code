<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Str;
use App\Job;
use App\Services\CategoryService;
use App\Services\JobService;
use App\Services\SubCategoryService;

class JobController extends Controller
{
    /**
    * @var JobService
    */
    protected $jobService;

    /**
    * @var CategoryService
    */
    protected $categoryService;

    /**
    * @var SubCategoryService
    */
    protected $subCategoryService;


    /**
    * Create a new controller instance.
    *
    * @return void
    */
    public function __construct(JobService $jService, CategoryService $cService, SubCategoryService $scService)
    {
        $this->middleware('auth', ['except' => ['show', 'showCategory', 'showSubCategory']]);

        $this->jobService = $jService;
        $this->categoryService = $cService;
        $this->subCategoryService = $scService;
    }

    /**
     * Show the form for creating a new resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function create()
    {
        $categories = $this->categoryService->getCategoriesCombo();
        $subcategories = [];
        $subcategories_select = [];

        return view('job.create')
            ->with('categories', $categories)
            ->with('subcategories', $subcategories)
            ->with('subcategories_select', $subcategories_select);
    }

    /**
     * Store a newly created resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return \Illuminate\Http\Response
     */
    public function store(Request $request)
    {
        $this->validateForm($request);

        $job = new Job;
        $this->saveData($job, $request);

        return redirect()->route('home');
    }

    /**
     * Display the specified resource.
     *
     * @param $id
     * @param $slug
     * @return \Illuminate\Http\Response
     */
    public function show($id, $slug)
    {
        $job = $this->jobService->getJob($id, $slug);
        return view('job.show')->with('job', $job);
    }

    /**
     * Show the form for editing the specified resource.
     *
     * @param  \App\Job  $job
     * @return \Illuminate\Http\Response
     */
    public function edit(Job $job)
    {
        $categories = $this->categoryService->getCategoriesCombo();
        $subcategories = $this->subCategoryService->getSubCategoriesCombo($job->category_id);
        $subcategories_select = $job->subcategories->pluck('id')->toArray();

        return view('job.edit')
            ->with('job', $job)
            ->with('categories', $categories)
            ->with('subcategories', $subcategories)
            ->with('subcategories_select', $subcategories_select);
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  \App\Job  $job
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, Job $job)
    {
        $this->validateForm($request);
        $this->saveData($job, $request);

        return redirect()->route('home');
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  \App\Job  $job
     * @return \Illuminate\Http\Response
     */
    public function destroy(Job $job)
    {
        Job::destroy($job->id);
        return redirect()->route('home');
    }

    /**
     * Get related jobs to category
     *
     * @param  $id Id category
     * @param  $slug slug category
     * @return \Illuminate\Http\Response
     */
    public function showCategory($id, $slug)
    {
        // Get category to get jobs
        $category = $this->categoryService->getCategory($id, $slug);

        // Get jobs
        $jobs = $this->jobService->getJobsCategory($id);

        return view('job.showCategory')
            ->with('category', $category)
            ->with('jobs', $jobs);
    }

    /**
     * Get related jobs to subcategory
     *
     * @param  $id Id category
     * @param  $slug slug category
     * @param  $subcategory_id Id subcategory
     * @param  $subcategory_slug slug subcategory
     * @return \Illuminate\Http\Response
     */
    public function showSubCategory($id, $slug, $subcategory_id, $subcategory_slug)
    {
        // Get subcategory to get jobs
        $subcategory = $this->subCategoryService->getSubCategories($slug, $subcategory_id, $subcategory_slug);

        // Get jobs
        $jobs = $this->jobService->getJobsSubCategory($id, $subcategory_id);

        return view('job.showSubCategory')
            ->with('subcategory', $subcategory)
            ->with('jobs', $jobs);
    }

    private function validateForm(Request $request)
    {
        $rules = [
            'title' => 'required',
            'content' => 'required',
            'contact' => 'required',
            'category_id' => 'required',
            'subcategory_id' => 'required'
        ];
        $niceNames = [
            'title' => 'Título',
            'content' => 'Contenido',
            'contact' => 'Contacto',
            'category_id' => 'Categoría',
            'subcategory_id' => 'Sub Categoría'
        ];

        $this->validate($request, $rules, [], $niceNames);
    }

    private function saveData($job, Request $request)
    {
        $job->title = $request->input('title');
        $job->slug = Str::slug($job->title, '-');
        $job->content = $request->input('content');
        $job->category_id = $request->input('category_id');
        $job->user_id = \Auth::user()->id;
        $job->contact = $request->input('contact');
        $job->save();

        $subcategory_id = $request->get('subcategory_id');
        $job->subcategories()->sync($subcategory_id);
    }
}