@extends('layouts.master')

@section('content')
<div class="container">
    <div class="row">
        <div class="col-md-12">
            <h1>{{ $file->title }}</h1>
            <iframe src="{{ asset($file->file_path) }}" style="width: 100%; height: 100vh;" frameborder="0"></iframe>
        </div>
    </div>
</div>
@endsection
