@extends('layouts.default')

@section('content')
<div class="panel panel-flat">
    <div class="panel-heading">
        <h5 class="panel-title">Edit Employee</h5>
    </div>
    <div class="panel-body">
        <form action="{{ route('employees.update', $employee->id) }}" method="POST">
            @csrf
            @method('PUT')
            <div class="form-group">
                <label for="tc_kimlik">TC Kimlik:</label>
                <input type="text" name="tc_kimlik" class="form-control" value="{{ $employee->tc_kimlik }}" required>
            </div>
            <div class="form-group">
                <label for="name">Name:</label>
                <input type="text" name="name" class="form-control" value="{{ $employee->name }}" required>
            </div>
            <div class="form-group">
                <label for="email">Email:</label>
                <input type="email" name="email" class="form-control" value="{{ $employee->email }}" required>
            </div>
            <div class="form-group">
                <label for="phone">Phone:</label>
                <input type="text" name="phone" class="form-control" value="{{ $employee->phone }}" required>
            </div>
            <div class="form-group">
                <label for="job_title">Job Title:</label>
                <input type="text" name="job_title" class="form-control" value="{{ $employee->job_title }}" required>
            </div>
            <div class="form-group">
                <label for="start_date">Start Date:</label>
                <input type="date" name="start_date" class="form-control" value="{{ $employee->start_date }}" required>
            </div>
            <button type="submit" class="btn btn-primary">Update</button>
        </form>
    </div>
</div>
@endsection
