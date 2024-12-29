@extends('layouts.master')

@section('content')
<div class="content-wrapper">
    <div class="page-header page-header-light">
        <div class="page-header-content header-elements-md-inline">
            <div class="page-title d-flex">
                <h4>
                    <i class="icon-arrow-left52 mr-2"></i>
                    <span class="font-weight-semibold">Prosedürler & Formlar</span>
                </h4>
                <a href="#" class="header-elements-toggle text-default d-md-none"><i class="icon-more"></i></a>
            </div>
        </div>

        <div class="breadcrumb-line breadcrumb-line-light header-elements-md-inline">
            <div class="d-flex">
                <div class="breadcrumb">
                    <a href="{{ route('admin.procedures-forms.index') }}" class="breadcrumb-item"><i
                            class="icon-home2 mr-2"></i> Prosedürler & Formlar</a>
                </div>
            </div>
        </div>
    </div>

    <div class="content">
        <div class="text-center mb-4">
            <h2 class="font-weight-bold">Prosedürler & Formlar</h2>
        </div>

        @if (session('success'))
        <div class="alert alert-success">
            {{ session('success') }}
        </div>
        @endif

        <!-- Arama Formu -->
        <form method="GET" action="{{ route('admin.procedures-forms.index') }}" class="mb-4">
            <div class="input-group">
                <input type="text" name="search" class="form-control" placeholder="Dosya ismi ile ara..." value="{{ request()->query('search') }}">
                <div class="input-group-append">
                    <button class="btn btn-primary" type="submit">Ara</button>
                </div>
            </div>
        </form>

        <!-- Dosyaların Listelenmesi -->
        <div class="card">
            <div class="card-body">
                <div class="row">
                    @foreach ($files as $file)
                    <div class="col-md-3">
                        <div class="card mb-4 shadow-sm">
                            <div class="card-body text-center">
                                <!-- Küçük simge ekleyin -->
                                <img src="{{ asset('images/formprosedür.png') }}" alt="File Icon" style="width: 50px; height: 50px;"/>
                                <h6 class="card-title mt-2">{{ $file->title }}</h6>
                                <p class="card-text">
                                    @if($file->user)
                                    Revize Eden: {{ $file->user->name }} {{ $file->user->surname }}
                                    @else
                                    Revize Eden: Bilinmiyor
                                    @endif
                                </p>
                                <a href="{{ route('admin.procedures-forms.download', $file->id) }}"
                                    class="btn btn-primary btn-sm">İndir</a>
                            </div>
                        </div>
                    </div>
                    @endforeach
                </div>
            </div>
        </div>

        <!-- Pagination Links -->
        {{ $files->links() }}

        <!-- Tümünü Gör Butonu -->
        @if(request()->query('search'))
        <div class="text-center mt-4">
            <a href="{{ route('admin.procedures-forms.index') }}" class="btn btn-secondary">Tümünü Gör</a>
        </div>
        @endif
    </div>
</div>
@endsection
