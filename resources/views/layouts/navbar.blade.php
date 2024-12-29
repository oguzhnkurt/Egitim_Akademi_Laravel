<!-- Main navbar -->
<nav class="navbar navbar-dark navbar-expand-lg navbar-static border-bottom border-bottom-white border-opacity-10">
    <div class="container-fluid">
        <!-- Mobile Toggle Button -->
        <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarContent"
            aria-controls="navbarContent" aria-expanded="false" aria-label="Toggle navigation">
            <span class="navbar-toggler-icon"></span>
        </button>

        <!-- Brand/Logo -->
        <div class="navbar-brand">
            <a href="{{ url('/index') }}" class="d-inline-flex align-items-center">
                <img src="{{ URL::asset('assets/images/logo_icon.svg') }}" alt="Logo" class="logo">
            </a>
        </div>

        <!-- Collapsible Navbar Content -->
        <div class="collapse navbar-collapse" id="navbarContent">
            <ul class="navbar-nav me-auto mb-2 mb-lg-0">
                <!-- Diğer Menü Öğeleri -->
                <li class="nav-item">
                    <a href="{{ url('/index') }}" class="nav-link">
                        <i class="ph-house"></i>
                        <span>Ana Sayfa</span>
                    </a>
                </li>

                <li class="nav-item">
                    <a href="{{ route('user.schedule.index') }}" class="nav-link">
                        <i class="ph-calendar"></i>
                        <span>Eğitim Takvimi</span>
                    </a>
                </li>
                <li class="nav-item">
                    <a href="{{ route('videos.index') }}" class="nav-link">
                        <i class="ph-book-open"></i>
                        <span>Akademik Eğitimler</span>
                    </a>
                </li>
                <li class="nav-item">
                    <a href="{{ route('videos.idari_egitimler') }}" class="nav-link">
                        <i class="ph-clipboard"></i>
                        <span>İdari Eğitimler</span>
                    </a>
                </li>
                <!-- Prosedürler & Formlar Dropdown -->
                <li class="nav-item dropdown">
                    <a href="#" class="nav-link dropdown-toggle white-text" data-bs-toggle="dropdown"
                        aria-expanded="false">
                        <i class="ph-file-text"></i>
                        <span style="color: #ffffff !important;">Prosedürler & Formlar</span>
                    </a>
                    <ul class="dropdown-menu dropdown-menu-end">
                        <!-- Prosedürler & Formlar Başlığı -->
                        <li style="color: #ffffff !important;"
                            class="dropdown-header text-uppercase fs-sm lh-sm sidebar-resize-hide white-text">
                            Prosedürler & Formlar
                        </li>
                        <li class="dropdown-divider"></li>

                        <!-- Akademi Bölümü -->
                        <li class="dropdown-submenu">
                            <a href="#" class="dropdown-item dropdown-toggle">Akademi</a>
                            <ul class="dropdown-menu">
                                <li><a href="{{ route('department.files', ['department' => 1, 'type' => 'prosedur']) }}"
                                        class="dropdown-item">Prosedürler</a></li>
                                <li><a href="{{ route('department.files', ['department' => 1, 'type' => 'is_akisi']) }}"
                                        class="dropdown-item">İş Akışı</a></li>
                                <li><a href="{{ route('department.files', ['department' => 1, 'type' => 'form']) }}"
                                        class="dropdown-item">Formlar</a></li>
                                <li><a href="{{ route('department.files', ['department' => 1, 'type' => 'dilekce']) }}"
                                        class="dropdown-item">Dilekçeler</a></li>
                            </ul>
                        </li>

                        <!-- Bilgi İşlem Bölümü -->
                        <li class="dropdown-submenu">
                            <a href="#" class="dropdown-item dropdown-toggle">Bilgi İşlem</a>
                            <ul class="dropdown-menu">
                                <li><a href="{{ route('department.files', ['department' => 2, 'type' => 'prosedur']) }}"
                                        class="dropdown-item">Prosedürler</a></li>
                                <li><a href="{{ route('department.files', ['department' => 2, 'type' => 'is_akisi']) }}"
                                        class="dropdown-item">İş Akışı</a></li>
                                <li><a href="{{ route('department.files', ['department' => 2, 'type' => 'form']) }}"
                                        class="dropdown-item">Formlar</a></li>
                                <li><a href="{{ route('department.files', ['department' => 2, 'type' => 'dilekce']) }}"
                                        class="dropdown-item">Dilekçeler</a></li>
                            </ul>
                        </li>

                        <!-- İhale Bölümü -->
                        <li class="dropdown-submenu">
                            <a href="#" class="dropdown-item dropdown-toggle">İhale</a>
                            <ul class="dropdown-menu">
                                <li><a href="{{ route('department.files', ['department' => 3, 'type' => 'prosedur']) }}"
                                        class="dropdown-item">Prosedürler</a></li>
                                <li><a href="{{ route('department.files', ['department' => 3, 'type' => 'is_akisi']) }}"
                                        class="dropdown-item">İş Akışı</a></li>
                                <li><a href="{{ route('department.files', ['department' => 3, 'type' => 'form']) }}"
                                        class="dropdown-item">Formlar</a></li>
                                <li><a href="{{ route('department.files', ['department' => 3, 'type' => 'dilekce']) }}"
                                        class="dropdown-item">Dilekçeler</a></li>
                            </ul>
                        </li>

                        <!-- İstatistik Bölümü -->
                        <li class="dropdown-submenu">
                            <a href="#" class="dropdown-item dropdown-toggle">İstatistik</a>
                            <ul class="dropdown-menu">
                                <li><a href="{{ route('department.files', ['department' => 4, 'type' => 'prosedur']) }}"
                                        class="dropdown-item">Prosedürler</a></li>
                                <li><a href="{{ route('department.files', ['department' => 4, 'type' => 'is_akisi']) }}"
                                        class="dropdown-item">İş Akışı</a></li>
                                <li><a href="{{ route('department.files', ['department' => 4, 'type' => 'form']) }}"
                                        class="dropdown-item">Formlar</a></li>
                                <li><a href="{{ route('department.files', ['department' => 4, 'type' => 'dilekce']) }}"
                                        class="dropdown-item">Dilekçeler</a></li>
                            </ul>
                        </li>

                        <!-- Finans Bölümü -->
                        <li class="dropdown-submenu">
                            <a href="#" class="dropdown-item dropdown-toggle">Finans</a>
                            <ul class="dropdown-menu">
                                <li><a href="{{ route('department.files', ['department' => 5, 'type' => 'prosedur']) }}"
                                        class="dropdown-item">Prosedürler</a></li>
                                <li><a href="{{ route('department.files', ['department' => 5, 'type' => 'is_akisi']) }}"
                                        class="dropdown-item">İş Akışı</a></li>
                                <li><a href="{{ route('department.files', ['department' => 5, 'type' => 'form']) }}"
                                        class="dropdown-item">Formlar</a></li>
                                <li><a href="{{ route('department.files', ['department' => 5, 'type' => 'dilekce']) }}"
                                        class="dropdown-item">Dilekçeler</a></li>
                            </ul>
                        </li>

                        <!-- Muhasebe Bölümü -->
                        <li class="dropdown-submenu">
                            <a href="#" class="dropdown-item dropdown-toggle">Muhasebe</a>
                            <ul class="dropdown-menu">
                                <li><a href="{{ route('department.files', ['department' => 6, 'type' => 'prosedur']) }}"
                                        class="dropdown-item">Prosedürler</a></li>
                                <li><a href="{{ route('department.files', ['department' => 6, 'type' => 'is_akisi']) }}"
                                        class="dropdown-item">İş Akışı</a></li>
                                <li><a href="{{ route('department.files', ['department' => 6, 'type' => 'form']) }}"
                                        class="dropdown-item">Formlar</a></li>
                                <li><a href="{{ route('department.files', ['department' => 6, 'type' => 'dilekce']) }}"
                                        class="dropdown-item">Dilekçeler</a></li>
                            </ul>
                        </li>

                        <!-- İnsan Kaynakları Bölümü -->
                        <li class="dropdown-submenu">
                            <a href="#" class="dropdown-item dropdown-toggle">İnsan Kaynakları</a>
                            <ul class="dropdown-menu">
                                <li><a href="{{ route('department.files', ['department' => 7, 'type' => 'prosedur']) }}"
                                        class="dropdown-item">Prosedürler</a></li>
                                <li><a href="{{ route('department.files', ['department' => 7, 'type' => 'is_akisi']) }}"
                                        class="dropdown-item">İş Akışı</a></li>
                                <li><a href="{{ route('department.files', ['department' => 7, 'type' => 'form']) }}"
                                        class="dropdown-item">Formlar</a></li>
                                <li><a href="{{ route('department.files', ['department' => 7, 'type' => 'dilekce']) }}"
                                        class="dropdown-item">Dilekçeler</a></li>
                            </ul>
                        </li>

                        <!-- Kalite Bölümü -->
                        <li class="dropdown-submenu">
                            <a href="#" class="dropdown-item dropdown-toggle">Kalite</a>
                            <ul class="dropdown-menu">
                                <li><a href="{{ route('department.files', ['department' => 8, 'type' => 'prosedur']) }}"
                                        class="dropdown-item">Prosedürler</a></li>
                                <li><a href="{{ route('department.files', ['department' => 8, 'type' => 'is_akisi']) }}"
                                        class="dropdown-item">İş Akışı</a></li>
                                <li><a href="{{ route('department.files', ['department' => 8, 'type' => 'form']) }}"
                                        class="dropdown-item">Formlar</a></li>
                                <li><a href="{{ route('department.files', ['department' => 8, 'type' => 'dilekce']) }}"
                                        class="dropdown-item">Dilekçeler</a></li>
                            </ul>
                        </li>
                    </ul>
                </li>
                <li class="nav-item">
                    <a href="{{ route('user.exam-portal.index') }}" class="nav-link">
                        <i class="ph-test-tube"></i>
                        <span>Sınav Portalı</span>
                    </a>
                </li>
            </ul>
        </div>


        <!-- Arama Çubuğu -->
        <form action="{{ route('search') }}" method="GET" class="d-flex align-items-center me-3 search-form">
            <input type="search" name="query" class="form-control rounded-pill search-input"
                placeholder="Ne arıyorsunuz?" aria-label="Search" aria-describedby="search-addon" />
            <button type="submit" class="btn btn-primary rounded-pill search-btn ms-2">Ara</button>
        </form>

        <!-- Kullanıcı Adı ve Logout Button -->
        <div class="d-flex align-items-center">
            <!-- Kullanıcı adı -->
            <span class="user-name me-3">Merhaba, {{ Auth::user()->name }}</span> <!-- Kullanıcı adını gösterir -->

            <!-- Logout Button -->
            <form id="logout-form" action="{{ route('logout') }}" method="POST" class="d-inline">
                @csrf
                @method('POST')
                <button type="submit" class="btn btn-custom rounded-pill px-4 py-2">
                    Çıkış Yap
                </button>
            </form>
        </div>
    </div>

    </div>
</nav>
<form id="logout-form" action="{{ route('logout') }}" method="POST" class="d-none">
    @csrf
    @method('POST')
</form>
<!-- /main navbar -->
<style>
    /* Siyah Renk için Kullanım */
    .black-text {
        color: #000000 !important;
    }

    /* Çıkış Butonu için Modern Stil */
    .btn-custom {
        border: 2px solid #ffffff;
        color: #ffffff;
        background-color: transparent;
        font-weight: 600;
        transition: all 0.3s ease-in-out;
        letter-spacing: 1px;
        text-transform: uppercase;
        border-radius: 50px;
    }

    /* Hover ve Focus Durumu */
    .btn-custom:hover {
        background-color: #ffffff;
        color: #343a40;
        border-color: #ffffff;
    }

    .btn-custom:focus,
    .btn-custom:active {
        outline: none;
        box-shadow: none;
    }

    /* Kullanıcı adı için stil */
    .user-name {
        font-size: 1rem;
        font-weight: 500;
        color: #ffffff;
    }

    /* Arama Çubuğu Stil */
    .search-form {
        display: flex;
        align-items: center;
    }

    .search-input {
        border: 2px solid #007bff;
        border-radius: 50px;
        padding-left: 20px;
        transition: all 0.3s ease;
        height: 2.5rem;
    }

    .search-input:focus {
        border-color: #0056b3;
        box-shadow: 0 0 10px rgba(0, 123, 255, 0.3);
        outline: none;
    }

    .search-btn {
        background-color: #007bff;
        color: white;
        border: none;
        padding-left: 20px;
        padding-right: 20px;
        height: 2.5rem;
        border-radius: 50px;
        transition: background-color 0.3s ease;
    }

    .search-btn:hover {
        background-color: #0056b3;
        color: white;
    }

    /* Mobil Görünüm */
    @media (max-width: 1080.98px) {
        .navbar-nav {
            margin-top: 1rem;
        }

        .navbar-nav .dropdown-menu {
            position: static;
            width: 100%;
            box-shadow: none;
            padding: 0;
            background-color: #495057;
        }

        .navbar-nav .dropdown-menu li {
            color: white;
        }

        .navbar-nav .dropdown-item {
            padding: 1rem;
            border-bottom: 1px solid white;
            color: white;
        }

        .navbar-nav .dropdown-item:hover {
            padding: 1rem;
            border-bottom: 1px solid rgb(0, 0, 0);
            color: rgb(0, 0, 0);
        }

        .navbar-nav .dropdown-item:last-child {
            border-bottom: none;
        }

        .search-form {
            width: 100%;
        }

        .search-input {
            width: 100%;
            margin-bottom: 10px;
        }
    }

    /* Tablet ve Üzeri İçin */
    @media (min-width: 768px) {
        .navbar-nav .dropdown-menu {
            position: absolute;
            top: 100%;
            left: 0;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
            padding: 0.5rem 0;
        }

        .navbar-nav .dropdown-menu .dropdown-item {
            padding: 0.75rem 1.25rem;
        }

        .search-form {
            max-width: 300px;
        }

        .search-input {
            width: 100%;
        }
    }

    /* Navbar brand (logo) boşluklarını kaldırmak için */
    .navbar-brand {
        padding: 0;
        /* Padding'i sıfırlıyoruz */
        margin: 0;
        /* Margin'i sıfırlıyoruz */
        display: flex;
        /* Flex kullanarak logo ve içeriği hizalayabiliriz */
        align-items: center;
        /* Dikey hizalama */
        max-width: 100px;
        /* Sınıf içeriği kadar büyüyüp küçülsün, içeriğindeki küçük olmasına rağmen kendisi büyük kalmasın */
    }

    /* Logo boyutunu optimize etmek isterseniz */
    .navbar-brand img {
        height: 32px;
        /* İstediğiniz yüksekliği ayarlayabilirsiniz */
        width: auto;
        /* Oranları koruyarak genişliği otomatik ayarlayın */
        margin-right: 10px;
        /* Logonun sağında bir miktar boşluk bırakabiliriz */
    }

    /* Menü öğelerini sola hizalamak için */
    .navbar-nav {
        margin-left: 0;
        /* Sol boşluğu sıfırlıyoruz */
        padding-left: 0;
        /* Sol boşluğu sıfırlıyoruz */
        flex-grow: 1;
        /* Menü öğelerinin tüm alanı kullanmasını sağlıyoruz */
        justify-content: flex-start;
        /* Menü öğelerini sola hizalıyoruz */
    }
</style>