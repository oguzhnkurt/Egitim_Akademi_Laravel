@extends('layouts.master')

@section('content')
<div class="container mt-5">
    <h2 class="text-center">Sınav Sonuçları</h2>

    <!-- Aylık Sonuçlar Butonu -->
    <div class="text-center mb-4">
        <a href="{{ route('exam.results.monthly') }}" class="btn btn-info">Aylık Sonuçlar</a>
    </div>

    <div class="card mt-4">
        <div class="card-header bg-primary text-white">
            <h5 class="mb-0">Sınav Sonuçları</h5>
        </div>

        <div class="card-body">
            <p>Aşağıda sınav sonuçları listelenmektedir. Başarılı olanlar yeşil, başarısız olanlar kırmızı olarak
                gösterilmektedir.</p>

            <!-- Excel ve PDF Görüntüleme Butonları -->
            <div class="text-center mb-4">
                <a href="{{ route('exam.results.excel') }}" class="btn btn-success">Excel Olarak İndir</a>
                <button class="btn btn-primary mb-3" onclick="openPDF()">PDF'i Görüntüle</button>
            </div>
        </div>

        <div class="table-responsive p-4">
            <table id="examResultsTable" class="table table-striped table-bordered">
                <thead class="table-dark">
                    <tr>
                        <th>Ad</th>
                        <th>Soyad</th>
                        <th>Görev</th>
                        <th>Bölge</th>
                        <th>Hastane</th>
                        <th>Durum</th>
                        <th>Puan</th>
                        <th>Ödül</th>
                    </tr>
                </thead>
                <tbody>
                    @foreach($results as $result)
                    <tr>
                        <td><a href="#" class="personel-link" data-id="{{ $result->user->id }}">{{ $result->user->name
                                }}</a></td>
                        <td>{{ $result->user->surname }}</td>
                        <td>{{ $result->user->job_title }}</td>
                        <td>{{ $result->user->region }}</td>
                        <td>{{ $result->user->hospital }}</td>
                        <td>
                            @if($result->score >= 70)
                            <span class="badge bg-success">Başarılı</span>
                            @else
                            <span class="badge bg-danger">Başarısız</span>
                            @endif
                        </td>
                        <td>{{ $result->score }}</td>
                        <td>{{ $result->reward }}</td>
                    </tr>
                    @endforeach
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- Geçmiş Sınav Sonuçları Modal -->
<div class="modal fade" id="examHistoryModal" tabindex="-1" aria-labelledby="examHistoryModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-lg">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="examHistoryModalLabel">Geçmiş Sınav Sonuçları</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Kapat"></button>
            </div>
            <div class="modal-body">
                <div class="table-responsive">
                    <table id="examHistoryTable" class="table table-striped table-bordered">
                        <thead class="table-dark">
                            <tr>
                                <th>Sınav Adı</th>
                                <th>Puan</th>
                                <th>Tarih</th>
                            </tr>
                        </thead>
                        <tbody id="examHistoryBody">
                            <!-- Geçmiş sınav sonuçları buraya eklenecek -->
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
</div>

@endsection

@section('scripts')
<link rel="stylesheet" href="https://cdn.datatables.net/1.13.4/css/jquery.dataTables.min.css">

<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>

<script>
    $(document).ready(function () {
        // Eğer tablo zaten DataTable ise, yok et
        if ($.fn.DataTable.isDataTable('#examResultsTable')) {
            $('#examResultsTable').DataTable().destroy();
        }

        // Yeni DataTable'ı başlat
        $('#examResultsTable').DataTable({
            pageLength: 25,
            order: [[0, 'asc']] // Ad sütununa göre sıralama
        });

        // Personel ismine tıklandığında modal açma
        $('.personel-link').on('click', function (e) {
            e.preventDefault();
            var userId = $(this).data('id');
            loadExamHistory(userId);
            $('#examHistoryModal').modal('show');
        });
    });

    // Geçmiş sınav sonuçlarını yükleyen fonksiyon
    function loadExamHistory(userId) {
        // AJAX isteği ile geçmiş sınav sonuçlarını alın
        $.ajax({
            url: '/get-exam-history/' + userId,
            method: 'GET',
            success: function (data) {
                $('#examHistoryBody').empty(); // Clear previous data
                if (Array.isArray(data)) {
                    data.forEach(function (result) {
                        $('#examHistoryBody').append(`
                            <tr>
                                <td>${result.exam_name}</td>
                                <td>${result.score}</td>
                                <td>${result.date}</td>
                            </tr>
                        `);
                    });
                } else {
                    console.error('Unexpected data format:', data);
                }
                // Reinitialize the DataTable
                $('#examHistoryTable').DataTable().destroy(); // Destroy previous instance
                $('#examHistoryTable').DataTable(); // Reinitialize
            },

            error: function () {
                alert('Geçmiş sınav sonuçları yüklenemedi.');
            }
        });
    }

    // PDF'i açma fonksiyonu
    function openPDF() {
        window.open("{{ route('exam.results.pdf') }}", "_blank");
    }
</script>
@endsection