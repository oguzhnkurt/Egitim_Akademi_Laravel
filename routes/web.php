<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\Admin\AdminScheduleController;
use App\Http\Controllers\HomeController;
use App\Http\Controllers\Auth\LoginController;
use App\Http\Controllers\CourseController;
use App\Http\Controllers\Admin\AdminVideoController;
use App\Http\Controllers\UserVideoController;
use App\Http\Controllers\EmployeeController;
use App\Http\Controllers\Admin\AdminExamController;
use App\Http\Controllers\TrainingController;
use App\Http\Controllers\MeetingController;
use App\Http\Controllers\Admin\AdminExamResultsController;
use App\Http\Controllers\TrainingAttendanceController;
use App\Http\Controllers\ReportController;
use App\Http\Controllers\MeetingAttendanceController;
use App\Http\Controllers\LimitlessController;
use App\Http\Controllers\UserExamController;
use App\Http\Controllers\ExamPortalController;
use App\Http\Controllers\ProceduresFormsController;
use App\Http\Controllers\FileController;
use App\Http\Controllers\SearchController;
use App\Http\Controllers\DepartmentController;

Route::get('/exam-results/excel', [AdminExamResultsController::class, 'exportExcel'])->name('exam.results.excel');
Route::get('/exam-results/pdf', [AdminExamResultsController::class, 'exportPDF'])->name('exam.results.pdf');
Route::get('/admin/exam-results/monthly', [AdminExamResultsController::class, 'monthlyResults'])->name('exam.results.monthly');
Route::get('/admin/exam-results/monthly/downloadExcel', [AdminExamResultsController::class, 'downloadExcel'])->name('admin.exam-results.monthly.downloadExcel');
Route::get('/exam-results/pdf', [AdminExamResultsController::class, 'exportPDF'])->name('exam.results.pdf');

Route::post('/exam-portal/finish/{exam}', [UserExamController::class, 'finishExam'])->name('user.exam-portal.finish');



// Kullanıcının sınav sonucunu gösteren rota
Route::get('/exam/{examId}/result', [AdminExamResultsController::class, 'showUserResult'])->name('exam.result.show');
Route::get('/exam/{examId}/user-result', [ExamPortalController::class, 'showUserResult'])
    ->name('user.exam.result');


route::get('/api/categories/{category}/subcategories', function ($categoryId) {
    $subcategories = Category::where('parent_id', $categoryId)->get();
    return response()->json($subcategories);
});

Route::get('/search', [SearchController::class, 'search'])->name('search');
Route::post('/admin/exams/upload-image', [AdminExamController::class, 'uploadImage'])->name('admin.exams.uploadImage');


// Dosya düzenleme
Route::put('/files/{id}', [FileController::class, 'update'])->name('files.update');
Route::get('/files/{id}/edit', [FileController::class, 'edit'])->middleware(['auth', 'role:admin'])->name('files.edit');

// Dosya görüntüleme ve indirme
Route::get('/files/open/{id}', [FileController::class, 'openFile'])->name('files.open');
Route::get('/files/view/{id}', [FileController::class, 'view'])->name('files.view');
Route::get('/files/download/{id}', [FileController::class, 'download'])->name('files.download');

// Dosya yükleme ve önizleme
Route::get('/files/upload', [FileController::class, 'showUploadForm'])->name('files.upload.form');
Route::post('/files/upload', [FileController::class, 'upload'])->name('files.upload');
Route::get('/files/preview/{id}', [FileController::class, 'viewPdf'])->name('files.preview');

// Dosyaları listeleme
Route::get('/files/{department}/{type}', [FileController::class, 'showFiles'])->name('files.index');

// Sadece belirli bir role sahip kullanıcılar için ek kısıtlamalar
Route::middleware('role:merkez-ofis')->group(function () {
    Route::get('/files/{department}/{type}', [FileController::class, 'showFiles'])->name('department.files');
    Route::get('/storage/files/{id}', [FileController::class, 'download'])->name('files.download');
});

Route::prefix('admin')->middleware(['auth', 'role:merkez-ofis'])->group(function () {
    Route::get('procedures-forms', [ProceduresFormsController::class, 'index'])->name('admin.procedures-forms.index');
    Route::post('procedures-forms', [ProceduresFormsController::class, 'store'])->name('admin.procedures-forms.store');
    Route::get('procedures-forms/download/{id}', [ProceduresFormsController::class, 'download'])->name('admin.procedures-forms.download');
});


// Admin Exam Routes
Route::prefix('admin')->middleware(['auth', 'role:admin'])->group(function () {
    Route::resource('exams', AdminExamController::class)->names([
        'index' => 'admin.exams.index',
        'create' => 'admin.exams.create',
        'store' => 'admin.exams.store',
        'show' => 'admin.exams.show',
        'edit' => 'admin.exams.edit',
        'update' => 'admin.exams.update',
        'destroy' => 'admin.exams.destroy',
    ]);
});
Route::get('/exam/{id}/end', [AdminExamController::class, 'end'])->name('exam.end');

Route::post('/exam-portal/start/{exam}', [ExamPortalController::class, 'start'])->name('user.exam-portal.start');

// User Exam Routes
Route::prefix('user')->middleware('auth')->group(function () {
    Route::get('exam-portal', [UserExamController::class, 'index'])->name('user.exam-portal.index');
    Route::post('/submit/{exam}', [UserExamController::class, 'submit'])->name('exam.submit');
    Route::get('exam-portal/{exam}', [UserExamController::class, 'start'])->name('user.exam-portal.start');
    Route::post('/exam-portal/start/{examId}', [ExamPortalController::class, 'start'])->name('user.exam-portal.start');
    Route::post('/exam/{exam}/submit', [ExamPortalController::class, 'submit'])->name('exam.submit');
});




// Admin Rotaları
Route::middleware(['auth', 'role:admin'])->group(function () {
    // Sınav ekleme formu
    Route::get('/admin/exams/create', [AdminExamController::class, 'create'])->name('admin.exams.create');

    // Sınavı kaydetme
    Route::post('/admin/exams/store', [AdminExamController::class, 'store'])->name('admin.exams.store');

    // Sınav sonuçlarını görüntüleme
    Route::get('/admin/exams/results', [AdminExamResultsController::class, 'results'])->name('admin.exams.results');
});


// Admin Exam Results Routes
Route::prefix('admin/exam-results')->middleware('admin')->group(function () {
    Route::get('/', [AdminExamResultsController::class, 'index'])->name('admin.exam-results.index');
    Route::get('/filter', [AdminExamResultsController::class, 'filter'])->name('admin.exam-results.filter');
    Route::get('/create', [AdminExamResultsController::class, 'create'])->name('admin.exam-results.create');
    Route::post('/', [AdminExamResultsController::class, 'store'])->name('admin.exam-results.store');
    Route::get('/{examResult}', [AdminExamResultsController::class, 'show'])->name('admin.exam-results.show');
    Route::get('/{examResult}/edit', [AdminExamResultsController::class, 'edit'])->name('admin.exam-results.edit');
    Route::put('/{examResult}', [AdminExamResultsController::class, 'update'])->name('admin.exam-results.update');
    Route::delete('/{examResult}', [AdminExamResultsController::class, 'destroy'])->name('admin.exam-results.destroy');
    Route::get('/exam/results/{id}', [AdminExamResultsController::class, 'show'])->name('exam.results');
    
    // Exam Results for specific exams
    Route::get('/exam/{exam}/results', [ExamPortalController::class, 'results'])->name('exam.results');
    Route::get('/get-exam-history/{userId}', [AdminExamResultsController::class, 'getExamHistory'])->name('admin.exam-results.history');
});

// Admin panel routes
Route::prefix('admin')->middleware(['auth', 'role:admin'])->group(function () {
    Route::get('schedule', [AdminScheduleController::class, 'index'])->name('schedule.index');
    Route::get('schedule/create', [AdminScheduleController::class, 'create'])->name('schedule.create');
    Route::post('/schedule/store', [AdminScheduleController::class, 'store'])->name('schedule.store');
    Route::get('schedule/{id}/edit', [AdminScheduleController::class, 'edit'])->name('schedule.edit');
    Route::put('schedule/{id}', [AdminScheduleController::class, 'update'])->name('schedule.update');
    Route::delete('schedule/{id}', [AdminScheduleController::class, 'destroy'])->name('schedule.destroy');
});

// Kullanıcılar için eğitim takvimi
Route::middleware('auth')->group(function () {
    Route::get('schedule', [AdminScheduleController::class, 'index'])->name('user.schedule.index');
});
// Kullanıcı rotaları
Route::middleware('auth')->group(function () {
    Route::get('users/roles', [App\Http\Controllers\UserController::class, 'showUsersWithRoles'])->name('users.roles');

    Route::get('users/assign-role', [App\Http\Controllers\UserController::class, 'showAssignRoleList'])
        ->middleware('admin')
        ->name('users.showAssignRoleList');

    // Kullanıcıya rol atama işlemi
    Route::post('users/assign-role', [App\Http\Controllers\UserController::class, 'assignRole'])
        ->middleware('admin')
        ->name('users.assignRole');
});
route::get('/admin/subcategories/{categoryId}', [AdminVideoController::class, 'getSubcategories']);
// Kullanıcı rotaları
Route::group(['middleware' => ['auth']], function () {
    Route::group(['middleware' => ['role:admin']], function () {
        Route::get('/dashboard', [AdminScheduleController::class, 'index'])->name('admin.dashboard');
        Route::get('/videos/create', [AdminVideoController::class, 'create'])->name('admin.videos.create');
        Route::get('/videos/{id}/edit', [AdminVideoController::class, 'edit'])->name('admin.videos.edit');
        Route::delete('/videos/{id}', [AdminVideoController::class, 'destroy'])->name('admin.videos.destroy');
    });

    Route::get('/videos', [UserVideoController::class, 'index'])->name('videos.index');
    Route::get('/videos/{id}', [UserVideoController::class, 'show'])->name('videos.show');
});
Route::middleware(['auth', 'admin'])->prefix('admin')->name('admin.')->group(function () {
    Route::resource('videos', AdminVideoController::class)->only(['store', 'destroy']);
});
// Admin rotaları
Route::prefix('admin')->middleware(['auth', 'admin'])->group(function () {
    Route::get('/videos', [AdminVideoController::class, 'index'])->name('admin.videos.index');
    Route::get('/videos/create', [AdminVideoController::class, 'create'])->name('admin.videos.create');
    Route::post('/videos', [AdminVideoController::class, 'upload'])->name('admin.videos.store');
    Route::get('/videos/{id}/edit', [AdminVideoController::class, 'edit'])->name('admin.videos.edit');
    Route::put('/videos/{id}', [AdminVideoController::class, 'update'])->name('admin.videos.update');
    Route::delete('/videos/{id}', [AdminVideoController::class, 'destroy'])->name('admin.videos.destroy');
});
Route::middleware('auth', 'role:admin')->group(function () {
    Route::get('/admin/idari-egitimler', [AdminVideoController::class, 'idariEgitimlerIndex'])->name('admin.videos.idari_egitimler');
    Route::post('/admin/idari-egitimler/upload', [AdminVideoController::class, 'upload'])->name('admin.videos.upload');
    Route::delete('/admin/idari-egitimler/{id}', [AdminVideoController::class, 'destroy'])->name('admin.videos.destroy');
});

// Kullanıcı Rotaları
route::get('/idari-egitimler', [UserVideoController::class, 'idariEgitimler'])->name('videos.idari_egitimler');

// Home page route
Route::get('/', function () {
    return view('welcome');
})->name('homepage');

Auth::routes();

// Login routes
Route::get('/login', [LoginController::class, 'showLoginForm'])->name('login');
Route::post('/login', [LoginController::class, 'login']);
Route::post('/logout', [LoginController::class, 'logout'])->name('logout');

// Redirect to home after login
Route::get('/home', [HomeController::class, 'home'])->middleware('auth')->name('home');

Route::get('/courses', [CourseController::class, 'index'])->name('courses.index');

Route::resource('employees', EmployeeController::class);
Route::resource('trainings', TrainingController::class);
Route::resource('trainingAttendance', TrainingAttendanceController::class);
Route::resource('reports', ReportController::class);
Route::resource('meetings', MeetingController::class);
Route::resource('meetingAttendance', MeetingAttendanceController::class);

// Authenticated user routes
Route::middleware(['auth'])->group(function () {
    // Catch-all route for any undefined route under authenticated users
    Route::get('/{any}', [LimitlessController::class, 'index'])->where('any', '.*')->name('catchall');
    
});
