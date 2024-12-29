@extends('layouts.master')

@section('content')

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
    transition: background-color 0.3s ease;
}

.dropdown-content {
    display: none;
    position: absolute;
    background-color: #f9f9f9;
    min-width: 160px;
    box-shadow: 0px 8px 16px 0px rgba(0, 0, 0, 0.2);
    z-index: 1;
    border-radius: 4px;
}

.dropdown-content a {
    color: black;
    padding: 12px 16px;
    text-decoration: none;
    display: block;
    border-bottom: 1px solid #ddd;
}

.dropdown-content a:hover {
    background-color: #f1f1f1;
}

.dropdown:hover .dropdown-content {
    display: block;
}

.dropdown:hover .dropbtn {
    background-color: #3e8e41;
    color: white;
}

.video-item {
    margin: 10px 0;
    padding: 10px;
    border: 1px solid #ddd;
    border-radius: 4px;
    background-color: #f9f9f9;
}

.video-item h5 {
    margin: 0 0 5px 0;
    font-weight: bold;
}

.video-item p {
    margin: 0;
    color: #666;
}

</style>


<div class="content-wrapper">
    <div class="page-header page-header-light">
        <div class="page-header-content header-elements-md-inline">
            <div class="page-title d-flex">
                <h4>
                    <i class="icon-arrow-left52 mr-2"></i>
                    <span class="font-weight-semibold">İdari Eğitimler</span>
                </h4>
                <a href="#" class="header-elements-toggle text-default d-md-none"><i class="icon-more"></i></a>
            </div>
        </div>

        <div class="breadcrumb-line breadcrumb-line-light header-elements-md-inline">
            <div class="d-flex">
                <div class="breadcrumb">
                    <a href="{{ route('videos.idari_egitimler') }}" class="breadcrumb-item"><i class="icon-home2 mr-2"></i> İdari Eğitimler</a>
                    <span class="breadcrumb-item active">Videolar</span>
                </div>
            </div>
        </div>
    </div>

    <div class="content">
        <div class="text-center mb-4">
            <h2 class="font-weight-bold">İdari Eğitim Videoları</h2>
        </div>

        <div class="card">
            <div class="card-body">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Başlık</th>
                            <th>İzle</th>
                        </tr>
                    </thead>
                    <tbody>
                        @foreach ($videos as $video)
                        @if($video->type === 'İdari Eğitim')
                        <tr>
                            <td>{{ $video->title }}</td>
                            <td>
                                <a href="{{ Storage::url($video->filename) }}" class="btn btn-success" target="_blank">İzle</a>
                            </td>
                        </tr>
                        @endif
                        @endforeach
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</div>
@endsection
<script>
    $(document).on('click', '.dropdown-item', function (e) {
    e.preventDefault();
    const subcategoryId = $(this).data('subcategory-id');

    // Fetch ile alt kategorilere göre videoları yükle
    fetch(`/videos?subcategory_id=${subcategoryId}`)
        .then(response => response.json())
        .then(data => {
            const videoList = document.getElementById('video-list');
            videoList.innerHTML = ''; // Önceki videoları temizler

            // Gelen videoları listele
            data.forEach(video => {
                const videoItem = document.createElement('div');
                videoItem.classList.add('video-item');
                videoItem.innerHTML = `
                    <h5>${video.title}</h5>
                    <p>${video.instructor.name}</p>
                    <video width="320" height="240" controls>
                        <source src="${video.filename}" type="video/mp4">
                        Tarayıcınız video etiketini desteklemiyor.
                    </video>
                `;
                videoList.appendChild(videoItem);
            });
        })
        .catch(error => {
            console.error('Videolar yüklenirken bir hata oluştu:', error);
        });
});

</script>
