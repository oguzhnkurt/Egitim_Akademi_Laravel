@extends('layouts.app')

@section('content')
    <h1>Eğitim Takvimi</h1>
    <table class="table">
        <thead>
            <tr>
                <th>Gün</th>
                <th>Saat</th>
                <th>Link</th>
                <th>Eğitmen</th>
            </tr>
        </thead>
        <tbody>
            @foreach ($schedules as $schedule)
                <tr>
                    <td>{{ $schedule->day }}</td>
                    <td>{{ $schedule->time }}</td>
                    <td><a href="{{ $schedule->link }}" target="_blank">Eğitime Katıl</a></td>
                    <td>{{ $schedule->instructor->name }}</td> <!-- Eğitmen ismini göster -->
                </tr>
            @endforeach
        </tbody>
    </table>
@endsection
