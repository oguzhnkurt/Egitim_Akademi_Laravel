@extends('layouts.master')

@section('content')
@component('components.breadcrumb')
    @slot('title') Sezin Akademi @endslot
    @slot('subtitle') İdari Eğitim Videoları @endslot
@endcomponent

<style>
    .btn-group .btn-primary {
        background-color: #007bff;
        border: none;
        padding: 8px 15px; /* Butonları küçülttüm */
        font-size: 14px;
        font-weight: bold;
        border-radius: 5px;
        transition: all 0.3s ease;
        margin-right: 5px; /* Marginleri küçülttüm */
    }

    .btn-group .btn-primary:hover {
        background-color: #0056b3;
        color: #fff;
        transform: scale(1.05);
    }

    .input-group {
        display: flex;
        width: 100%;
        max-width: 600px;
        margin: 0 auto;
    }

    .input-group .form-control {
        border-top-right-radius: 0;
        border-bottom-right-radius: 0;
        padding: 8px 15px; /* Arama kutusu ve butonları eşitledim */
        font-size: 14px;
        flex: 1;
    }

    .input-group .input-group-append .btn-primary {
        background-color: #28a745;
        border: none;
        font-size: 14px;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 8px 15px; /* Arama butonunu küçülttüm */
    }

    .input-group .input-group-append .btn-primary:hover {
        background-color: #218838;
        transform: scale(1.05);
    }

    .card.shadow-sm {
        transition: all 0.3s ease;
    }

    .card.shadow-sm:hover {
        box-shadow: 0 8px 16px rgba(0, 0, 0, 0.15);
        transform: translateY(-5px);
    }

    .card .card-body h5 {
        font-size: 18px;
        color: #007bff;
        font-weight: bold;
    }

    .card .btn-primary {
        background-color: #007bff;
        border: none;
    }

    .card .btn-primary:hover {
        background-color: #0056b3;
    }

    #video-container {
        background-color: #fff;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
    }
</style>

<div class="content-wrapper">
    <div class="content">
        <div class="text-center mb-4">
            <h2 class="title-effect">İdari Eğitim Videoları</h2>
        </div>

        <div class="d-flex justify-content-center mb-4">
            @foreach($categories as $category)
            <div class="btn-group mr-2">
                <button type="button" class="btn btn-primary dropdown-toggle" data-bs-toggle="dropdown"
                    aria-expanded="false">
                    {{ $category->name }}
                </button>
                <ul class="dropdown-menu">
                    @foreach($category->subcategories as $subcategory)
                    <li><a class="dropdown-item subcategory-btn" href="#" data-category-id="{{ $category->id }}"
                            data-subcategory-id="{{ $subcategory->id }}">
                            {{ $subcategory->name }}
                        </a></li>
                    @endforeach
                </ul>
            </div>
            @endforeach
        </div>

        <div class="d-flex justify-content-center mb-4">
            <form method="GET" action="{{ route('videos.index') }}" class="form-inline">
                <div class="input-group">
                    <input type="text" id="search" name="search" class="form-control" placeholder="Videoları Ara..."
                        value="{{ request()->input('search') }}">
                    <div class="input-group-append">
                        <button type="submit" class="btn btn-primary">Ara</button>
                    </div>
                </div>
            </form>
        </div>

        <div id="video-container" class="card mb-4">
            <div class="card-body">
                <div class="row">
                    @forelse ($videos as $video)
                    <div class="col-md-2 mb-4">
                        <div class="card shadow-sm">
                            <a href="{{ route('videos.show', $video->id) }}">
                                <img src="{{ asset('images/akademi.png') }}" class="card-img-top"
                                    alt="{{ $video->title }}">
                            </a>
                            <div class="card-body">
                                <h5 class="card-title">
                                    <a href="{{ route('videos.show', $video->id) }}"
                                        class="font-weight-bold text-primary">{{ $video->title }}</a>
                                </h5>
                                <p class="card-text">Eğitmen: {{ $video->instructor->name }} {{
                                    $video->instructor->surname }}</p>
                                <p class="card-text">Eğitim Tarihi: {{ $video->date->translatedFormat('d F Y') }}</p>
                                <a href="{{ route('videos.show', $video->id) }}" class="btn btn-primary">İzlemeye
                                    Başla</a>
                            </div>
                        </div>
                    </div>
                    @empty
                    <p class="text-center">Bu kategoriye ait video bulunamadı.</p>
                    @endforelse
                </div>
            </div>
        </div>
    </div>
</div>

<script>
    document.addEventListener('DOMContentLoaded', function () {
        document.querySelectorAll('.subcategory-btn').forEach(function (button) {
            button.addEventListener('click', function (e) {
                e.preventDefault();
                const categoryId = this.getAttribute('data-category-id');
                const subcategoryId = this.getAttribute('data-subcategory-id');

                let url = new URL(window.location.href);
                url.searchParams.set('category_id', categoryId);
                url.searchParams.set('subcategory_id', subcategoryId);

                window.location.href = url.toString();
            });
        });
    });
</script>

@endsection
