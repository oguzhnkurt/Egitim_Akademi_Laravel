@extends('layouts.master')

@section('content')
<style>
    .badge-custom {
        background-color: black;
        color: white;
    }
</style>

<div class="container">
    <h1>Kullanıcı Rolleri</h1>
    <table class="table table-bordered mt-3">
        <thead>
            <tr>
                <th>ID</th>
                <th>İsim</th>
                <th>E-posta</th>
                <th>Roller</th>
            </tr>
        </thead>
        <tbody>
            @foreach($users as $user)
            <tr>
                <td>{{ $user->id }}</td>
                <td>{{ $user->name }}</td>
                <td>{{ $user->email }}</td>
                <td>
                    @foreach($user->roles as $role)
                        <span class="badge badge-custom">{{ $role->name }}</span>
                    @endforeach
                </td>
                
            </tr>
            @endforeach
        </tbody>
    </table>
</div>
@endsection
