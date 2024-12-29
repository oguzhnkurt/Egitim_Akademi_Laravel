@extends('layouts.master')

@section('content')
    @component('components.breadcrumb')
        @slot('title') Sezin Akademi @endslot
        @slot('subtitle') Yeni Eğitim Oluştur @endslot
    @endcomponent

    <!-- Content area -->
    <div class="content">

        <!-- Create Schedule -->
        <div class="card">
            <div class="card-header bg-primary text-white">
                <h5 class="mb-0">Yeni Eğitim Ekle</h5>
            </div>

            <div class="card-body">
                <form method="POST" action="{{ route('schedule.store') }}">
                    @csrf

                    <div class="mb-3">
                        <label for="day" class="form-label">Gün</label>
                        <input type="text" name="day" class="form-control" id="day" placeholder="Örneğin: Pazartesi" required>
                    </div>

                    <div class="mb-3">
                        <label for="time" class="form-label">Saat</label>
                        <input type="text" name="time" class="form-control" id="time" placeholder="Örneğin: 10:00 - 11:00" required>
                    </div>

                    <div class="mb-3">
                        <label for="link" class="form-label">Link</label>
                        <input type="text" name="link" class="form-control" id="link" placeholder="Zoom linkini giriniz" required>
                    </div>

                    <!-- Eğitmen Seç -->
                    <div class="mb-3">
                        <label for="instructor_selection" class="form-label">Eğitmen Seçimi</label>
                        <select name="instructor_selection" id="instructor_selection" class="form-control" required>
                            <option value="" disabled selected>Varolan Eğitmenler</option>
                            @foreach($instructors as $instructor)
                                <option value="existing-{{ $instructor->user->id }}">
                                    {{ $instructor->name }} {{ $instructor->surname }} 
                                    ({{ $instructor->external ? 'Dış Eğitmen' : 'İç Eğitmen' }})
                                </option>
                            @endforeach
                            <option value="new">Yeni Eğitmen Ekle</option>
                        </select>
                        
                    </div>

                    <!-- Yeni Eğitmen Bilgileri -->
                    <div id="new-instructor-fields" style="display: none;">
                        <div class="mb-3">
                            <label for="instructor_name" class="form-label">Ad</label>
                            <input type="text" name="instructor_name" class="form-control" id="instructor_name" placeholder="Eğitmenin adı">
                        </div>
                        <div class="mb-3">
                            <label for="instructor_surname" class="form-label">Soyad</label>
                            <input type="text" name="instructor_surname" class="form-control" id="instructor_surname" placeholder="Eğitmenin soyadı">
                        </div>
                        <div class="mb-3">
                            <label for="instructor_email" class="form-label">E-posta</label>
                            <input type="email" name="instructor_email" class="form-control" id="instructor_email" placeholder="Eğitmenin e-postası">
                        </div>
                        <div class="mb-3">
                            <label for="instructor_type" class="form-label">Eğitmen Türü</label>
                            <select name="instructor_type" id="instructor_type" class="form-control">
                                <option value="" disabled selected>Eğitmen Türü Seçin</option>
                                <option value="internal">İç Eğitmen</option>
                                <option value="external">Dış Eğitmen</option>
                            </select>
                        </div>
                    </div>

                    <div class="mt-3 d-flex justify-content-end">
                        <button type="submit" class="btn btn-success">Kaydet</button>
                    </div>
                </form>
            </div>
        </div>
        <!-- /Create Schedule -->

    </div>
    <!-- /content area -->
@endsection

@section('scripts')
<script>
    document.addEventListener('DOMContentLoaded', function () {
        var instructorSelection = document.getElementById('instructor_selection');
        var newInstructorFields = document.getElementById('new-instructor-fields');

        if (instructorSelection) {
            instructorSelection.addEventListener('change', function () {
                if (this.value === 'new') {
                    newInstructorFields.style.display = 'block';
                } else {
                    newInstructorFields.style.display = 'none';
                }
            });
        }
    });
</script>
@endsection
