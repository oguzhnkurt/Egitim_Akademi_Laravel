@extends('layouts.master')

@section('content')
    @component('components.breadcrumb')
        @slot('title') Sezin Akademi @endslot
        @slot('subtitle') Eğitim Takvimi @endslot
    @endcomponent

    <!-- Content area -->
    <div class="content">

        <!-- Eğitim Takvimi -->
        <div class="card">
            <div class="card-header">
                <h5 class="mb-0">Eğitim Tarihleri</h5>
                @if(auth()->user()->hasRole('admin'))
                    <a href="{{ route('schedule.create') }}" class="btn btn-primary float-right">Yeni Eğitim Ekle</a>
                @endif
            </div>

            <div class="card-body">
                <div class="mb-4">
                    Sezin Akademi, öğrenmeyi zaman ve mekanla sınırlı tutmayan, eğitimin her zaman,
                    her yerde, devamlı olmasının önemini biliyor. İnovasyon ve Gelişim Akademisi çatısı altında tüm personellerimizin gelişimine hem örgün hem uzaktan eğitimlerimizle değer katıyoruz.
                    <br><br>
                    <code>ÖNEMLİ UYARI: Eğitim linki dersten 5 dakika önce "Link" kısmında gözükecektir.</code>
                </div>

                <div class="table-responsive">
                    <table class="table table-bordered">
                        <thead>
                            <tr>
                                <th>Gün</th>
                                <th>Saat</th>
                                <th>Link</th>
                                <th>Eğitmen</th>
                                @if(auth()->user()->hasRole('admin'))
                                    <th>Eylemler</th>
                                @endif
                            </tr>
                        </thead>
                        <tbody>
                            @foreach ($schedules as $schedule)
                                <tr>
                                    <td>{{ $schedule->day }}</td>
                                    <td>{{ $schedule->time }}</td>
                                    <td><a href="{{ $schedule->link }}" target="_blank">Linke Git</a></td>
                                    <td>{{ $schedule->instructor->name }} {{ $schedule->instructor->surname }}</td>
                                    @if(auth()->user()->hasRole('admin'))
                                        <td>
                                            <a href="{{ route('schedule.edit', $schedule->id) }}" class="btn btn-warning">Düzenle</a>
                                            <form action="{{ route('schedule.destroy', $schedule->id) }}" method="POST" style="display:inline;">
                                                @csrf
                                                @method('DELETE')
                                                <button type="submit" class="btn btn-danger">Sil</button>
                                            </form>
                                        </td>
                                    @endif
                                </tr>
                            @endforeach
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        <!-- /Eğitim Takvimi -->

    </div>
    <!-- /content area -->
@endsection
