@extends('layouts.master')

<!-- jQuery -->
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>

<!-- Bootstrap JavaScript -->
<script src="https://stackpath.bootstrapcdn.com/bootstrap/5.3.0/js/bootstrap.bundle.min.js"></script>

<!-- DataTables CSS -->
<link rel="stylesheet" href="https://cdn.datatables.net/1.13.4/css/jquery.dataTables.min.css">

<!-- DataTables JS -->
<script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>

<!-- DataTables Buttons CSS -->
<link rel="stylesheet" href="https://cdn.datatables.net/buttons/1.7.2/css/buttons.dataTables.min.css">

<!-- DataTables Buttons JS -->
<script src="https://cdn.datatables.net/buttons/1.7.2/js/dataTables.buttons.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.1.3/jszip.min.js"></script>
<script src="https://cdn.datatables.net/buttons/1.7.2/js/buttons.html5.min.js"></script>
<script src="https://cdn.datatables.net/buttons/1.7.2/js/buttons.print.min.js"></script>
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">

@section('content')

<!-- Video ekleme bölümü -->
<div class="video-container">
    <video id="backgroundVideo" autoplay muted playsinline>
        <source src="{{ asset('images/prosedürler.mp4') }}" type="video/mp4">
        Tarayıcınız bu video formatını desteklemiyor.
    </video>
</div>

<div class="container mt-4">
    <div class="row">
        <div class="col-md-12">
            <h1 class="mb-4">Dosyalar</h1>
            <div class="mb-3">
                <a href="{{ route('files.upload.form') }}" class="btn btn-primary">
                    <i class="fas fa-upload"></i> Yeni Dosya Yükle
                </a>
            </div>
            <div class="table-responsive"> <!-- Tabloyu kaydırılabilir yapıyoruz -->
                <table id="filesTable" class="table table-hover table-bordered">
                    <thead class="thead-dark">
                        <tr>
                            <th>Başlık</th>
                            <th>Departman</th>
                            <th>Tür</th>
                            <th>İşlemler</th>
                        </tr>
                    </thead>
                    <tbody>
                        @forelse ($files as $file)
                        <tr>
                            <td>{{ $file->title }}</td>
                            <td>{{ $file->department->name }}</td>
                            <td>{{ ucfirst($file->type) }}</td>
                            <td>
                                <!-- Prosedür Dosyalarını Yalnızca Görüntüle -->
                                @if($file->type === 'prosedur')
                                <a href="{{ route('files.open', ['id' => $file->id]) }}" class="btn btn-info btn-sm"
                                    target="_blank">
                                    <i class="fas fa-eye"></i> Görüntüle
                                </a>
                                @else
                                <!-- Diğer Dosya Türleri için Görüntüle, İndir ve Düzenle -->
                                <a href="{{ route('files.open', ['id' => $file->id]) }}" class="btn btn-info btn-sm"
                                    target="_blank">
                                    <i class="fas fa-eye"></i> Görüntüle
                                </a>
                                <a href="{{ route('files.download', ['id' => $file->id]) }}"
                                    class="btn btn-success btn-sm">
                                    <i class="fas fa-download"></i> İndir
                                </a>

                                <!-- Düzenle Butonu (Sadece dilekçe ve form dosyaları için) -->
                                @if($file->type === 'dilekce' || $file->type === 'form')
                                <a href="{{ route('files.edit', ['id' => $file->id]) }}" class="btn btn-warning btn-sm">
                                    <i class="fas fa-edit"></i> Düzenle
                                </a>
                                @endif
                                @endif
                            </td>

                        </tr>
                        @empty
                        <tr>
                            <td colspan="4" class="text-center">Henüz dosya bulunmamaktadır.</td>
                        </tr>
                        @endforelse
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</div>
@endsection



@push('scripts')
<script>
    $(document).ready(function () {
        // DataTable kurulumu
        $('#filesTable').DataTable({
            "language": {
                "url": "//cdn.datatables.net/plug-ins/1.10.25/i18n/Turkish.json"
            },
            "dom": 'Bfrtip',
            "buttons": [
                'copy', 'csv', 'excel', 'pdf', 'print'
            ]
        });

        // Video bittiğinde, videoyu gizle
        $('#backgroundVideo').on('ended', function () {
            $(this).fadeOut();
        });
    });
</script>
@endpush

<style>
    .video-container {
        position: relative;
        width: 100%;
        height: 25vh;
        /* Videonun yüksekliği sabit ve çok büyük olmayacak */
        overflow: hidden;
        background-color: black;
        /* Video yüklenmediğinde arka plan rengi */
        z-index: -1;
        /* Video'nun navbar'ın altında kalmasını sağlıyoruz */
    }

    .video-container video {
        position: absolute;
        top: 50%;
        left: 50%;
        width: 100%;
        height: 100%;
        object-fit: cover;
        /* Videonun tam sığması için */
        transform: translate(-50%, -50%);
    }

    /* Navbar stili - gereksiz uzamaları engellemek için sadeleştirildi */
    .navbar {
        position: relative;
        z-index: 1000;
        /* Navbar video üstünde kalacak */
        background-color: rgba(255, 255, 255, 0.9);
        /* Şeffaf arka plan */
        padding: 10px 20px;
        /* Yatay ve dikey padding ayarları */
        display: flex;
        justify-content: space-between;
        /* İçeriklerin sağ ve sol hizalanmasını sağlar */
        align-items: center;
        /* Dikey hizalama */
        height: auto;
        /* Gereksiz yükseklik uzamasını engeller */
    }

    .navbar .search-container {
        display: flex;
        align-items: center;
    }

    .navbar .search-container input[type="text"] {
        height: 40px;
        margin-right: 10px;
        padding: 5px 10px;
        border-radius: 5px;
        border: 1px solid #ccc;
    }

    .navbar .search-container button {
        height: 40px;
        padding: 5px 15px;
        background-color: #007bff;
        color: white;
        border: none;
        border-radius: 5px;
        cursor: pointer;
    }

    .navbar .search-container button:hover {
        background-color: #0056b3;
    }

    /* Mobil uyumlu düzen */
    @media (max-width: 768px) {
        .video-container {
            height: 20vh;
            /* Küçük ekranlarda video yüksekliği daha da küçültüldü */
        }

        .navbar .search-container input[type="text"] {
            width: 100%;
            /* Küçük ekranlarda arama kutusunu tam genişlik yapma */
        }
    }
</style>