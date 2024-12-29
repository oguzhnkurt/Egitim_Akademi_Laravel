@extends('layouts.master')

@section('content')
    <div class="content-wrapper">
        <div class="page-header page-header-light">
            <div class="page-header-content header-elements-md-inline">
                <div class="page-title d-flex">
                    <h4><i class="icon-arrow-left52 mr-2"></i> <span class="font-weight-semibold">Video Detayı</span></h4>
                    <a href="#" class="header-elements-toggle text-default d-md-none"><i class="icon-more"></i></a>
                </div>
            </div>

            <div class="breadcrumb-line breadcrumb-line-light header-elements-md-inline">
                <div class="d-flex">
                    <div class="breadcrumb">
                        <a href="{{ route('videos.index') }}" class="breadcrumb-item"><i class="icon-home2 mr-2"></i> Akademik Eğitimler</a>
                        <span class="breadcrumb-item active">Video Detayı</span>
                    </div>
                </div>
            </div>
        </div>

        <div class="content">
            <!-- Kapatılabilir Uyarı Bulutu -->
            <div class="d-flex justify-content-center">
                <div class="alert alert-success alert-dismissible fade show text-center position-relative" role="alert" style="width: 50%;">
                    <button type="button" class="btn-close position-absolute top-0 end-0 mt-2 me-2" data-bs-dismiss="alert" aria-label="Close"></button>
                    <h4 class="alert-heading">Dikkat!</h4>
                    <p>Eğitim videolarımız yazılım ekibimizin geliştirdiği özel bir yazılım sayesinde ileriye sardırılamamaktadır.</p>
                    <hr>
                    <p class="mb-0">Eğitim videolarını dikkatle izleyerek en iyi şekilde yararlanabilirsiniz. İleri sarma özelliği devre dışıdır.</p>
                </div>
            </div>

            <div class="card">
                <div class="card-header header-elements-inline">
                    <h5 class="card-title">{{ $video->title }}</h5>
                </div>

                <div class="card-body">
                    <video width="100%" height="auto" controls controlsList="nodownload" >
                        <source src="{{ asset('storage/videos/' . $video->filename) }}" type="video/mp4">
                        Tarayıcınız video etiketini desteklemiyor.
                    </video>
                    <p class="mt-3">{{ $video->description }}</p>
                </div>
            </div>
        </div>
    </div>
@endsection

<style>@media (max-width: 768px) {
    .alert {
        width: 100%;
    }
}
</style>
