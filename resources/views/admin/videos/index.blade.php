@extends('layouts.master')

@section('content')
<div class="content-wrapper">
    <div class="page-header page-header-light">
        <div class="page-header-content header-elements-md-inline">
            <div class="page-title d-flex">
                <h4><i class="icon-arrow-left52 mr-2"></i> <span class="font-weight-semibold">Videolar</span></h4>
                <a href="#" class="header-elements-toggle text-default d-md-none"><i class="icon-more"></i></a>
            </div>
        </div>

        <div class="breadcrumb-line breadcrumb-line-light header-elements-md-inline">
            <div class="d-flex">
                <div class="breadcrumb">
                    <a href="{{ route('admin.dashboard') }}" class="breadcrumb-item"><i
                            class="icon-home2 mr-2"></i>Anasayfa</a>
                    <span class="breadcrumb-item active">Videolar</span>
                </div>
            </div>
        </div>
    </div>

    <div class="content">
        <div class="card">
            <div class="card-header header-elements-inline">
                <h5 class="card-title">Video Listesi</h5>
                <div class="header-elements">
                    <a href="{{ route('admin.videos.create') }}" class="btn btn-primary btn-labeled btn-labeled-right">
                        <b><i class="icon-plus3"></i></b> Yeni Video Yükle
                    </a>
                </div>
            </div>
            <div class="d-flex justify-content-center mb-4">
                <!-- Ana Kategori Butonları -->
                <div class="btn-group" role="group" aria-label="Kategori Seçimi">
                    @foreach($categories as $category)
                    <div class="btn-group">
                        <button type="button" class="btn btn-secondary dropdown-toggle" data-toggle="dropdown"
                            aria-haspopup="true" aria-expanded="false">
                            {{ $category->name }}
                        </button>
                        <div class="dropdown-menu">
                            <!-- Alt Kategoriler -->
                            @foreach($category->children as $subcategory)
                            <a class="dropdown-item" href="#" data-subcategory-id="{{ $subcategory->id }}">
                                {{ $subcategory->name }}
                            </a>
                            @endforeach
                        </div>
                    </div>
                    @endforeach
                </div>
            </div>
            <!-- Kategori Seçimi ve Arama Formu -->
            <div class="card-body">
                <form method="GET" action="{{ route('admin.videos.index') }}" class="form-inline mb-3">
                    <div class="input-group mr-2">
                        <select id="category_id" name="category_id" class="form-control">
                            <option value="">Kategori Seçin</option>
                            @foreach ($categories as $category)
                            <option value="{{ $category->id }}" {{ request()->input('category_id') == $category->id ?
                                'selected' : '' }}>
                                {{ $category->name }}
                            </option>
                            @endforeach
                        </select>
                    </div>
                    <div class="input-group">
                        <input type="text" id="search" name="search" class="form-control" placeholder="Videoları Ara..."
                            value="{{ request()->input('search') }}">
                        <div class="input-group-append">
                            <button type="submit" class="btn btn-primary">Ara</button>
                        </div>
                    </div>
                </form>
            </div>

            <div class="table-responsive">
                <table class="table datatable-basic">
                    <thead>
                        <tr>
                            <th>ID</th>
                            <th>Başlık</th>
                            <th>Video</th>
                            <th class="text-center">İşlemler</th>
                        </tr>
                    </thead>
                    <tbody>
                        @foreach ($videos as $video)
                        <tr>
                            <td>{{ $video->id }}</td>
                            <td>{{ $video->title }}</td>
                            <td>
                                <video width="320" height="240" controls>
                                    <source src="{{ Storage::url($video->filename) }}" type="video/mp4">
                                    Tarayıcınız video etiketini desteklemiyor.
                                </video>
                            </td>
                            <td class="text-center">
                                <div class="list-icons">
                                    <div class="dropdown">
                                        <button class="dropbtn">
                                            <i class="icon-menu9"></i>
                                        </button>
                                        <div class="dropdown-content">
                                            <a href="{{ route('admin.videos.edit', $video->id) }}">
                                                <i class="icon-pencil7"></i> Düzenle
                                            </a>
                                            <form action="{{ route('admin.videos.destroy', $video->id) }}" method="POST"
                                                style="display: inline;">
                                                @csrf
                                                @method('DELETE')
                                                <button type="submit" class="dropdown-item">
                                                    <i class="icon-trash"></i> Sil
                                                </button>
                                            </form>
                                        </div>
                                    </div>
                                </div>
                            </td>
                        </tr>
                        @endforeach
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</div>
@endsection

@section('scripts')
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://cdn.datatables.net/1.10.24/js/jquery.dataTables.min.js"></script>
<script>
    $(document).on('click', '.dropdown-item', function (e) {
        e.preventDefault();
        const subcategoryId = $(this).data('subcategory-id');

        // Fetch ile alt kategorilere göre videoları yükle
        fetch(`/videos?subcategory_id=${subcategoryId}`)
            .then(response => response.json())
            .then(data => {
                const videoList = document.getElementById('video-list');
                videoList.innerHTML = '';

                // Gelen videoları listele
                data.forEach(video => {
                    const videoItem = document.createElement('div');
                    videoItem.classList.add('video-item');
                    videoItem.innerHTML = `<h5>${video.title}</h5><p>${video.instructor.name}</p>`;
                    videoList.appendChild(videoItem);
                });
            })
            .catch(error => {
                console.error('Videolar yüklenirken bir hata oluştu:', error);
            });
    });
</script>
@endsection

<style>
    .dropdown {
        position: relative;
        display: inline-block;
    }

    .dropbtn {
        background-color: #f9f9f9;
        color: black;
        padding: 12px 16px;
        font-size: 16px;
        border: none;
        cursor: pointer;
        border-radius: 4px;
    }

    .dropdown-content {
        display: none;
        position: absolute;
        background-color: #f9f9f9;
        min-width: 160px;
        box-shadow: 0px 8px 16px 0px rgba(0, 0, 0, 0.2);
        z-index: 1;
    }

    .dropdown-content a {
        color: black;
        padding: 12px 16px;
        text-decoration: none;
        display: block;
    }

    .dropdown-content a:hover {
        background-color: #f1f1f1
    }

    .dropdown:hover .dropdown-content {
        display: block;
    }

    .dropdown:hover .dropbtn {
        background-color: #3e8e41;
    }
</style>