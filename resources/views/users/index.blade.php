<!-- resources/views/users/index.blade.php -->
@extends('layouts.master')

@section('content')
<div class="content-wrapper">
    <div class="page-header page-header-light">
        <div class="page-header-content header-elements-md-inline">
            <div class="page-title d-flex">
                <h4>
                    <i class="icon-arrow-left52 mr-2"></i>
                    <span class="font-weight-semibold">Kullanıcı Listesi</span>
                </h4>
                <a href="#" class="header-elements-toggle text-default d-md-none"><i class="icon-more"></i></a>
            </div>
        </div>
    </div>

    <div class="content">
        <div class="card">
            <div class="card-body">
                @if (session('success'))
                    <div class="alert alert-success">
                        {{ session('success') }}
                    </div>
                @endif
                <table class="table">
                    <thead>
                        <tr>
                            <th>Adı</th>
                            <th>Soyadı</th>
                            <th>İşlem</th>
                        </tr>
                    </thead>
                    <tbody>
                        @foreach ($users as $user)
                        <tr>
                            <td>{{ $user->name }}</td>
                            <td>{{ $user->surname }}</td>
                            <td>
                                <a href="{{ route('users.showAssignRoleForm', $user->id) }}" class="btn btn-primary">Rol Ata</a>
                            </td>
                        </tr>
                        @endforeach
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</div>
@endsection

@section('center-scripts')
<!-- Theme JS files -->
<script src="{{URL::asset('assets/js/vendor/visualization/d3/d3.min.js')}}"></script>
<script src="{{URL::asset('assets/js/vendor/visualization/d3/d3_tooltip.js')}}"></script>
@endsection

@section('scripts')
<script src="{{URL::asset('assets/demo/pages/dashboard.js')}}"></script>
<script src="{{URL::asset('assets/demo/charts/pages/dashboard/streamgraph.js')}}"></script>
<script src="{{URL::asset('assets/demo/charts/pages/dashboard/sparklines.js')}}"></script>
<script src="{{URL::asset('assets/demo/charts/pages/dashboard/lines.js')}}"></script>
<script src="{{URL::asset('assets/demo/charts/pages/dashboard/areas.js')}}"></script>
<script src="{{URL::asset('assets/demo/charts/pages/dashboard/donuts.js')}}"></script>
<script src="{{URL::asset('assets/demo/charts/pages/dashboard/bars.js')}}"></script>
<script src="{{URL::asset('assets/demo/charts/pages/dashboard/progress.js')}}"></script>
<script src="{{URL::asset('assets/demo/charts/pages/dashboard/heatmaps.js')}}"></script>
<script src="{{URL::asset('assets/demo/charts/pages/dashboard/pies.js')}}"></script>
<script src="{{URL::asset('assets/demo/charts/pages/dashboard/bullets.js')}}"></script>
@endsection
