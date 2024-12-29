@extends('layouts.master')

@section('content')
<div class="background-container">
    <video autoplay muted loop>
        <source src="{{ asset('images/exampagebackground.mp4') }}" type="video/mp4">
        Tarayıcınız video formatını desteklemiyor.
    </video>

    <div class="container content mt-4">
        <h1 class="mb-4 text-center">Mevcut Sınavlar</h1> <!-- Başlık ortaya alındı -->
        <div class="table-responsive">
            <table class="table table-striped table-bordered">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Başlık</th>
                        <th>Açıklama</th>
                        <th>Başlangıç Tarihi</th>
                        <th>Bitiş Tarihi</th>
                        <th>Süre</th>
                        <th>İşlemler</th>
                    </tr>
                </thead>
                <tbody>
                    @foreach($exams as $exam)
                    <tr id="exam-row-{{ $exam->id }}">
                        <td>{{ $exam->id }}</td>
                        <td>{{ $exam->title }}</td>
                        <td>{{ $exam->description }}</td>
                        <td>{{ $exam->start_date }}</td>
                        <td>{{ $exam->end_date }}</td>
                        <td>{{ $exam->duration }}</td>
                        <td>
                            <button type="button" class="btn btn-primary" data-toggle="modal"
                                data-target="#warningModal" data-exam-id="{{ $exam->id }}">
                                Sınava Gir
                            </button>

                            <!-- Sınavı Bitir Butonu, sadece admin yetkisi olanlar görebilir -->
                            @can('admin')
                            <form method="POST" action="{{ route('user.exam-portal.finish', $exam->id) }}" class="d-inline-block finish-exam-form" id="finish-exam-form-{{ $exam->id }}">
                                @csrf
                                <button type="submit" class="btn btn-danger finish-exam-button" data-exam-id="{{ $exam->id }}">Sınavı Bitir</button>
                            </form>
                            @endcan
                        </td>
                    </tr>

                    <!-- Her sınav için 15 dakika sonra satırı kaldıracak zamanlayıcı -->
                    <script>
                        setTimeout(function() {
                            document.getElementById('exam-row-{{ $exam->id }}').style.display = 'none';
                        }, 900000); // 15 dakika (900000 ms)
                    </script>
                    @endforeach
                </tbody>
            </table>
        </div>

        <!-- Sadece admin kullanıcılar için Sınav Sonuçlarını Göster ve Yeni Sınav Oluştur Butonları -->
        @if(auth()->check() && auth()->user()->hasRole('admin'))
        <div class="text-center mt-5">
            <!-- Sınav Sonuçlarını Göster Butonu -->
            <button type="button" class="btn btn-secondary" data-toggle="modal" data-target="#resultModal">
                Sınav Sonuçlarını Göster
            </button>

            <!-- Yeni Sınav Oluştur Butonu -->
            <a href="{{ route('admin.exams.create') }}" class="btn btn-success ml-3">
                Yeni Sınav Oluştur
            </a>
        </div>
        @endif
    </div>
</div>

<!-- Uyarı Modal -->
<div class="modal fade" id="warningModal" tabindex="-1" role="dialog" aria-labelledby="warningModalLabel"
    aria-hidden="true">
    <div class="modal-dialog modal-dialog-scrollable" role="document">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="warningModalLabel">Sınav Uyarısı</h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <div class="modal-body">
                Sınava başlamadan önce lütfen aşağıdaki uyarıları dikkatlice okuyun:<br><br>
                1. Sınav başladığında ekranınız tam ekran olacak ve pencereniz küçültülemeyecek.<br>
                2. Sınav sırasında sekme değiştiremeyeceksiniz ve tarayıcıyı kapatmanız durumunda sınavınız geçersiz
                sayılacaktır.<br>
                3. Sınav sırasında sorulara geri dönme hakkınız olmayacak ve tüm soruları sırasıyla yanıtlamanız
                gerekecek.<br><br>
                <input type="checkbox" id="agreeCheckbox"> Uyarıları okudum ve anladım.
            </div>
            <div class="modal-footer">
                <form id="startExamForm" method="POST">
                    @csrf
                    <button type="button" class="btn btn-secondary" data-dismiss="modal">Kapat</button>
                    <button type="submit" class="btn btn-primary" id="startExamButton" disabled>Sınava Başla</button>
                </form>
            </div>
        </div>
    </div>
</div>

<!-- Sınav Sonuçları Modal -->
<div class="modal fade" id="resultModal" tabindex="-1" role="dialog" aria-labelledby="resultModalLabel"
    aria-hidden="true">
    <div class="modal-dialog modal-dialog-scrollable" role="document">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="resultModalLabel">Sınav Sonuçlarını Görüntüle</h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <div class="modal-body">
                <form id="viewResultForm" method="GET" action="{{ route('admin.exam-results.index') }}">
                    <div class="form-group">
                        <label for="examSelect">Hangi sınavın sonucunu görmek istiyorsunuz?</label>
                        <select class="form-control" id="examSelect" name="examId">
                            @foreach($exams as $exam)
                            <option value="{{ $exam->id }}">{{ $exam->title }}</option>
                            @endforeach
                        </select>
                    </div>
                </form>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-dismiss="modal">Kapat</button>
                <button type="submit" form="viewResultForm" class="btn btn-primary">Sonuçları Görüntüle</button>
            </div>
        </div>
    </div>
</div>

@endsection

@section('scripts')
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@4.6.0/dist/js/bootstrap.bundle.min.js"></script>

<script>
    $(document).ready(function () {
        $('#warningModal').on('show.bs.modal', function (event) {
            var button = $(event.relatedTarget);
            var examId = button.data('exam-id');
            var form = $(this).find('form');
            var actionUrl = "{{ route('user.exam-portal.start', '') }}/" + examId;
            form.attr('action', actionUrl);
        });

        $('#agreeCheckbox').on('change', function () {
            $('#startExamButton').prop('disabled', !this.checked);
        });

        $('#startExamButton').on('click', function () {
            $('#startExamForm').submit();
        });

        // Sınavı Bitir butonuna basıldığında sınavı bitir ve butonu kaybolsun
        $('.finish-exam-button').on('click', function(e) {
            e.preventDefault();
            var button = $(this);
            var examId = button.data('exam-id');

            // Formu gönder
            $('#finish-exam-form-' + examId).submit();

            // 15 dakika (900000 ms) sonra sınav satırını gizle
            setTimeout(function() {
                $('#exam-row-' + examId).fadeOut('slow');
            }, 900000); // 15 dakika
        });
    });
</script>

<style>
    /* Video arka planı */
    .background-container {
        position: relative;
        width: 100%;
        height: 100vh;
        overflow: hidden;
    }

    .background-container video {
        position: absolute;
        top: 50%;
        left: 50%;
        min-width: 100%;
        min-height: 100%;
        width: auto;
        height: auto;
        z-index: -1;
        transform: translate(-50%, -50%);
        background-size: cover;
    }

    /* İçerik kutusu */
    .content {
        position: relative;
        z-index: 2;
        padding: 20px;
        background-color: rgba(255, 255, 255, 0.92);
        border-radius: 10px;
    }

    .table-responsive {
        background-color: rgba(255, 255, 255, 0.3);
    }

    /* Tablo yazıları için stil */
    .table td, .table th {
        color: black; /* Tablo içindeki yazıları siyah yapıyoruz */
    }

    /* Başlık ortalanmış */
    h1.text-center {
        text-align: center;
    }

    /* Buton stilleri */
    .text-center {
        margin-top: 30px;
    }

    .btn-secondary {
        margin-top: 15px;
    }
    
    .btn-success {
        margin-top: 15px;
    }
</style>
@endsection

@section('nofooter')
@endsection
