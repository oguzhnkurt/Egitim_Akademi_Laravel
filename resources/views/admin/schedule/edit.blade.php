@extends('layouts.master')

@section('content')
    @component('components.breadcrumb')
        @slot('title') Sezin Akademi @endslot
        @slot('subtitle') Eğitimi Düzenle @endslot
    @endcomponent

    <!-- Content area -->
    <div class="content">

        <!-- Edit Schedule -->
        <div class="card">
            <div class="card-header">
                <h5 class="mb-0">Eğitimi Düzenle</h5>
            </div>

            <div class="card-body">
                <!-- Başarı Mesajı -->
                @if(session('success'))
                    <div class="alert alert-success">
                        {{ session('success') }}
                    </div>
                @endif

                <form method="POST" action="{{ route('schedule.update', $schedule->id) }}">
                    @csrf
                    @method('PUT')

                    <!-- Gün ve Saat Input Alanları -->
                    <div class="form-group">
                        <label for="day">Gün</label>
                        <input type="text" name="day" class="form-control" id="day" value="{{ $schedule->day }}" required>
                    </div>

                    <div class="form-group">
                        <label for="time">Saat</label>
                        <input type="text" name="time" class="form-control" id="time" value="{{ $schedule->time }}" required>
                    </div>

                    <div class="form-group">
                        <label for="link">Link</label>
                        <input type="text" name="link" class="form-control" id="link" value="{{ $schedule->link }}" required>
                    </div>

                    <div class="form-group">
                        <label for="instructor_id">Eğitmen Seçin</label>
                        <select name="instructor_id" id="instructor_id" class="form-control" onchange="toggleNewInstructorFields(this)">
                            <option value="" disabled selected>Eğitmen Seçin</option>
                            <option value="new_instructor">Yeni Eğitmen Ekle</option>
                            @foreach ($instructors as $instructor)
                                <option value="{{ $instructor->id }}" {{ $instructor->id == $schedule->instructor_id ? 'selected' : '' }}>
                                    {{ $instructor->name }} {{ $instructor->surname }}
                                </option>
                            @endforeach
                        </select>
                    </div>

                    <div id="newInstructorFields" style="display: none;">
                        <div class="form-group">
                            <label for="new_instructor_type">Eğitmen Türü</label>
                            <select name="new_instructor_type" id="new_instructor_type" class="form-control">
                                <option value="internal">İç Eğitmen</option>
                                <option value="external">Dış Eğitmen</option>
                            </select>
                        </div>

                        <div class="form-group">
                            <label for="new_instructor_name">Yeni Eğitmen Adı</label>
                            <input type="text" name="new_instructor_name" class="form-control" id="new_instructor_name" placeholder="Eğitmen adı">
                        </div>

                        <div class="form-group">
                            <label for="new_instructor_surname">Yeni Eğitmen Soyadı</label>
                            <input type="text" name="new_instructor_surname" class="form-control" id="new_instructor_surname" placeholder="Eğitmen soyadı">
                        </div>

                        <div class="form-group">
                            <label for="new_instructor_email">Yeni Eğitmen Email</label>
                            <input type="email" name="new_instructor_email" class="form-control" id="new_instructor_email" placeholder="Eğitmen emaili">
                        </div>
                    </div>

                    <div class="mt-3">
                        <button type="submit" class="btn btn-primary">Güncelle</button>
                    </div>
                </form>
            </div>
        </div>
        <!-- /Edit Schedule -->

    </div>
    <!-- /content area -->

    <script>
        function toggleNewInstructorFields(selectElement) {
            const newInstructorFields = document.getElementById('newInstructorFields');
            if (selectElement.value === 'new_instructor') {
                newInstructorFields.style.display = 'block';
            } else {
                newInstructorFields.style.display = 'none';
            }
        }
    </script>
@endsection
