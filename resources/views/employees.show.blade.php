@extends('layouts.default')

@section('content')
<div class="panel panel-flat">
    <div class="panel-heading">
        <h5 class="panel-title">Employee Details</h5>
    </div>
    <div class="panel-body">
        <div class="form-group">
            <label>TC Kimlik:</label>
            <p>{{ $employee->tc_kimlik }}</p>
        </div>
        <div class="form-group">
            <label>Name:</label>
            <p>{{ $employee->name }}</p>
        </div>
        <div class="form-group">
            <label>Email:</label>
            <p>{{ $employee->email }}</p>
        </div>
        <div class="form-group">
            <label>Phone:</label>
            <p>{{ $employee->phone }}</p>
        </div>
        <div class="form-group">
            <label>Job Title:</label>
            <p>{{ $employee->job_title }}</p>
        </div>
        <div class="form-group">
            <label>Start Date:</label>
            <p>{{ $employee->start_date }}</p>
        </div>
        <a href="{{ route('employees.index') }}" class="btn btn-default">Back</a>
    </div>
</div>
@endsection
