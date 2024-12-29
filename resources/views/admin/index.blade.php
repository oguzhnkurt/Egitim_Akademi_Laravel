@extends('layouts.master')

@section('content')

<!-- Ana Sayfa İçerik Başlangıcı -->
<div class="container mt-5">
    <!-- Tıklanabilir Küçük Görsel -->
    <div class="row justify-content-center mb-5">
        <div class="col-md-6 text-center">
            <a href="{{ url('/schedule') }}">
                <img src="{{ asset('images/2024.png') }}" alt="Sezin Tıbbi Görüntüleme"
                    class="img-fluid clickable-image">
            </a>
        </div>
    </div>

    <!-- Videolar: İki Video Yan Yana Alt Köşelere Hizalı -->
    <div class="row mb-5">
        <!-- Video 1: Akademik Eğitim Videoları -->
        <div class="col-md-6 text-left">
            <a href="{{ url('/videos') }}" class="video-link">
                <video autoplay muted loop class="img-fluid rounded shadow">
                    <source src="{{ asset('images/egitimvideo.mp4') }}" type="video/mp4">
                    Tarayıcınız bu videoyu oynatamıyor.
                </video>
                <h3 class="mt-3 text-center">Akademik Eğitim Videoları</h3>
            </a>
        </div>

        <!-- Video 2: İdari Eğitim Videoları -->
        <div class="col-md-6 text-right">
            <a href="{{ url('/idari-egitimler') }}" class="video-link">
                <video autoplay muted loop class="img-fluid rounded shadow">
                    <source src="{{ asset('images/idariegitim.mp4') }}" type="video/mp4">
                    Tarayıcınız bu videoyu oynatamıyor.
                </video>
                <h3 class="mt-3 text-center">İdari Eğitim Videoları</h3>
            </a>
        </div>
    </div>

    <!-- Video: MR Odaları Videosu -->
    <div class="row justify-content-center mb-5">
        <div class="col-md-8 text-center">
            <video autoplay muted loop class="img-fluid rounded shadow">
                <source src="{{ asset('images/mrvideoları.mp4') }}" type="video/mp4">
                Tarayıcınız bu videoyu oynatamıyor.
            </video>
        </div>
    </div>

    <!-- Bilgilendirici Kartlar -->
    <div class="container mt-5">
        <div class="container mt-5">
            <div class="row justify-content-center">
                <div class="col-md-10">
                    <!-- Ana Başlık ve Metin Kartı -->
                    <div class="card service-card text-center p-4">
                        <h2 class="fw-bold title">Sezin Tıbbi Görüntüleme ve Kalp Merkezi</h2>
                        <p class="lead text-muted">
                            İnsanların huzurlu ve mutlu bir şekilde yaşam sürmeleri için sağlıkları önemlidir. <br>
                            Sağlığınız için tüm hijyenik koşulları sağlıyor, doğru tedavi için doğru teşhisin önemini
                            bilerek, <br>
                            sağlığınıza kavuşabilmeniz için inovatif teknolojiye yatırım yapıyoruz.
                        </p>
                    </div>
                </div>
            </div>
        </div>


        <div class="row text-center justify-content-center">
            <!-- Hijyenik Kartı -->
            <div class="col-md-4 mb-4">
                <div class="card service-card h-100">
                    <div class="card-body">
                        <div class="icon-box">
                            <i class="fas fa-pump-medical"></i>
                        </div>
                        <h3 class="card-title">Hijyenik</h3>
                        <p class="card-text">MR & CT odalarımızın hijyenik olmasına son derece önem veriyoruz. Düzenli
                            olarak hasta değişimlerinde dezenfekte işlemleri gerçekleştiriyoruz.</p>
                    </div>
                </div>
            </div>

            <!-- Akademik Kartı -->
            <div class="col-md-4 mb-4">
                <div class="card service-card h-100">
                    <div class="card-body ">
                        <div class="icon-box">
                            <i class="fas fa-graduation-cap"></i>
                        </div>
                        <h3 class="card-title">Akademik</h3>
                        <p class="card-text">Öğrenmeyi zaman ve mekâna sınırlı tutmayan, eğitimin her zaman, her yerde,
                            devamlı olmasının önemini biliyor, İnovasyon ve Gelişim Akademisi çatısı altında
                            eğitimlerimize devam ediyoruz.</p>
                    </div>
                </div>
            </div>

            <!-- Teknolojik Kartı -->
            <div class="col-md-4 mb-4">
                <div class="card service-card h-100">
                    <div class="card-body">
                        <div class="icon-box">
                            <i class="fas fa-microchip"></i>
                        </div>
                        <h3 class="card-title">Teknolojik</h3>
                        <p class="card-text">Tıbbi teknoloji, geniş bir sağlık ürünleri yelpazesini kapsıyor. Sezin
                            ailesi olarak inovatif teknolojiyi yakından takip ediyor, görüntüleme cihazlarımızın güncel
                            teknolojiyle uyum sağlamasını önemsiyor.</p>
                    </div>
                </div>
            </div>
        </div>
    </div>


    <!-- Eklenen Sınav Başarı Videosu -->
    <!-- Sabit Görsel: Başarı Mesajı -->
    <div class="row justify-content-center mb-5">
        <div class="col-md-6 text-center">
            <img src="{{ asset('images/bassarilar.png') }}" alt="Sınav Başarı Mesajı" class="img-fluid clickable-image">
        </div>
    </div>


    <!-- Bize Ulaşın Butonu: Sayfanın En Altında Sabit ve Şık -->
    <div class="row mt-5 justify-content-center align-items-center">
        <div class="col-md-12 text-center">
            <!-- Bize Ulaşın Butonu -->
            <button class="btn btn-outline-info" data-bs-toggle="modal" data-bs-target="#contactModal">
                <i class="fas fa-envelope"></i> Bize Ulaşın
            </button>
        </div>
    </div>
    <div style="height: 20px;"> </div>



    <!-- Modal: İletişim Bilgileri -->
    <div class="modal fade" id="contactModal" tabindex="-1" aria-labelledby="contactModalLabel" aria-hidden="true">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="contactModalLabel">Bizimle İletişime Geçin!</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <p>Her biri kendi alanında uzman ekibimiz ile sizlere yardımcı olmaya hazırız.</p>
                    <div class="d-flex justify-content-between">
                        <div>
                            <h6><i class="fas fa-map-marker-alt"></i> Adres</h6>
                            <p>Fevziçakmak Mahallesi<br>Bukas Ticaret Merkezi<br>10576. Sk. No:1/104,
                                42050<br>Karatay/Konya</p>
                        </div>
                        <div>
                            <h6><i class="fas fa-phone"></i> Telefon</h6>
                            <p>0 (332) 323 33 06<br>+90 505 064 6941</p>
                        </div>
                    </div>
                    <div class="mt-3">
                        <h6><i class="fas fa-envelope"></i> Mail</h6>
                        <p>oguzhan.kurt@sezintip.com</p>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Kapat</button>
                </div>
            </div>
        </div>
    </div>


</div>

@endsection

<!-- CSS -->
@push('styles')
<style>
    /* Genel Tasarım */
    .container {
        margin-top: 50px;
        max-width: 1200px;
    }

    .contact-btn {
        background-color: #FF6F00;
        /* Ana turuncu renk */
        color: white;
        border: none;
        padding: 15px 40px;
        font-size: 1.2rem;
        font-weight: 600;
        text-transform: uppercase;
        border-radius: 50px;
        cursor: pointer;
        transition: background-color 0.3s ease, transform 0.2s ease, box-shadow 0.3s ease;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.15);
        /* Hafif gölge */
    }

    .contact-btn i {
        margin-right: 10px;
        /* İkon ile metin arasında boşluk */
        font-size: 1.5rem;
    }

    /* Hover Efekti */
    .contact-btn:hover {
        background-color: #FF8C00;
        /* Daha açık turuncu */
        transform: translateY(-5px);
        /* Hafif yukarı kalkma */
        box-shadow: 0 8px 16px rgba(0, 0, 0, 0.2);
        /* Gölgenin genişlemesi */
    }

    /* Buton Aktif (Tıklanma Durumu) */
    .contact-btn:active {
        transform: scale(0.98);
        /* Hafif küçültme efekti */
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        /* Hafif gölge */
    }

    /* Başlık ve Metin */
    .title {
        font-size: 2.5rem;
        font-weight: bold;
        color: #2c3e50;
        margin-bottom: 20px;
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    .lead {
        font-size: 1.2rem;
        line-height: 1.6;
        color: #7f8c8d;
        margin-bottom: 40px;
    }

    /* Kart Genel Tasarımı */
    .service-card {
        background: #fff;
        border: none;
        border-radius: 15px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        transition: transform 0.3s ease, box-shadow 0.3s ease;
        padding: 30px;
    }

    .service-card:hover {
        transform: translateY(-10px);
        box-shadow: 0 10px 20px rgba(0, 0, 0, 0.15);
    }

    /* Kart Başlıkları */
    .card-title {
        font-size: 1.5rem;
        font-weight: 700;
        color: #34495e;
        margin-top: 15px;
        text-transform: uppercase;
    }

    /* Kart Metni */
    .card-text {
        font-size: 1rem;
        color: #7f8c8d;
        margin-top: 10px;
    }

    /* İkonlar */
    .icon-box {
        font-size: 3rem;
        color: #16a085;
        margin-bottom: 15px;
    }

    /* Video Ayarları */
    .video-link video,
    video {
        max-width: 100%;
        height: auto;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
    }

    .video-link video:hover {
        transform: scale(1.03);
        transition: transform 0.3s ease-in-out;
    }

    /* Responsive Tasarım */
    @media (max-width: 768px) {
        .title {
            font-size: 2rem;
        }

        .lead {
            font-size: 1rem;
        }

        .icon-box {
            font-size: 2.5rem;
        }

        .card-title {
            font-size: 1.3rem;
        }

        .card-text {
            font-size: 0.9rem;
        }

        .btn-contact {
            font-size: 1rem;
            padding: 12px 30px;
        }
    }
</style>
@endpush

<!-- Bootstrap JavaScript -->
@push('scripts')
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
@endpush