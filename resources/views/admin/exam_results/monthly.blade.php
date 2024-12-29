@extends('layouts.master')

@section('content')
<div class="container mt-5">
    <h2 class="text-center">Aylık Sınav Sonuçları</h2>
    <div class="card mt-4">
        <div class="card-header bg-info text-white">
            <h5 class="mb-0">Bu Ayın Sınav Sonuçları</h5>
        </div>

        <div class="card-body">
            <p>Aşağıda bu ay içinde girilen sınavların sonuçları ve kazanılan toplam ödüller listelenmektedir.</p>
        </div>
        <div class="table-responsive p-4">
            <a href="{{ route('admin.exam-results.monthly.downloadExcel') }}" class="btn btn-success mb-3">Excel Olarak
                İndir</a>
            <table id="monthlyResultsTable" class="table table-striped table-bordered">
                <thead class="table-dark">
                    <tr>
                        <th>Kullanıcı Adı</th>
                        <th>Görev</th>
                        <th>Bölge</th>
                        <th>Hastane</th>
                        <th>Toplam Ödül</th>
                        <th>Girilen Sınavlar ve Puanlar</th>
                    </tr>
                </thead>
                <tbody>
                    @foreach($monthlyRewards as $reward)
                    <tr>
                        <td>{{ $reward['user']->name ?? 'N/A' }} {{ $reward['user']->surname ?? 'N/A' }}</td>
                        <td>{{ $reward['user']->job_title ?? 'N/A' }}</td>
                        <td>{{ $reward['user']->region ?? 'N/A' }}</td>
                        <td>{{ $reward['user']->hospital ?? 'N/A' }}</td>
                        <td>{{ $reward['total_reward'] }}</td>
                        <td>
                            <ul>
                                @foreach($reward['exams'] as $exam)
                                <li>{{ $exam->exam->title ?? 'Bilinmiyor' }}: {{ $exam->score ?? 0 }} puan</li>
                                @endforeach
                            </ul>
                        </td>
                    </tr>
                    @endforeach
                </tbody>
            </table>
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
        // DataTable daha önce başlatılmadıysa başlatılır
        if (!$.fn.DataTable.isDataTable('#monthlyResultsTable')) {
            $('#monthlyResultsTable').DataTable({
                pageLength: 25,
                order: [[0, 'asc']] // Kullanıcı adına göre sıralama
            });
        }
    });
</script>
@endsection