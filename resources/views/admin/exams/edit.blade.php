@extends('layouts.master')

@section('content')
<div class="container">
    <h1>Sınavı Düzenle</h1>
    <form action="{{ route('admin.exams.update', $exam->id) }}" method="POST">
        @csrf
        @method('PUT')
        <div class="form-group">
            <label for="title">Başlık</label>
            <input type="text" class="form-control" id="title" name="title" value="{{ $exam->title }}" required>
        </div>
        <div class="form-group">
            <label for="description">Açıklama</label>
            <textarea class="form-control" id="description" name="description" required>{{ $exam->description }}</textarea>
        </div>
        <div class="form-group">
            <label for="start_date">Başlangıç Tarihi</label>
            <input type="date" class="form-control" id="start_date" name="start_date" value="{{ $exam->start_date }}" required>
        </div>
        <div class="form-group">
            <label for="end_date">Bitiş Tarihi</label>
            <input type="date" class="form-control" id="end_date" name="end_date" value="{{ $exam->end_date }}" required>
        </div>
        <div class="form-group">
            <label for="duration">Süre (dakika)</label>
            <input type="number" class="form-control" id="duration" name="duration" value="{{ $exam->duration }}" required>
        </div>

        <hr>
        <h3>Sorular</h3>
        <div id="questions-container">
            @foreach($exam->questions as $index => $question)
                <div class="question">
                    <h4>Soru {{ $index + 1 }}</h4>
                    <div class="form-group">
                        <label for="questions[{{ $index }}][question_text]">Soru Metni</label>
                        <textarea class="form-control" id="questions[{{ $index }}][question_text]" name="questions[{{ $index }}][question_text]" required>{{ $question->question_text }}</textarea>
                    </div>
                    <div class="form-group">
                        <label for="questions[{{ $index }}][max_score]">Maksimum Puan</label>
                        <input type="number" class="form-control" id="questions[{{ $index }}][max_score]" name="questions[{{ $index }}][max_score]" value="{{ $question->max_score }}" required>
                    </div>
                    <div class="form-group">
                        <label>Seçenekler (Doğru cevabı seçiniz)</label>
                        <div class="options-container">
                            @foreach(json_decode($question->options, true) as $letter => $text)
                                <div class="option">
                                    <input type="radio" id="question_{{ $index }}_option_{{ $letter }}" name="questions[{{ $index }}][correct_answer]" value="{{ $letter }}" {{ $question->correct_answer === $letter ? 'checked' : '' }} required>
                                    <label for="question_{{ $index }}_option_{{ $letter }}">Seçenek {{ $letter }}</label>
                                    <input type="text" class="form-control" name="questions[{{ $index }}][options][]" value="{{ $text }}" required>
                                </div>
                            @endforeach
                        </div>
                        <button type="button" class="btn btn-secondary add-option" data-question-index="{{ $index }}">Seçenek Ekle</button>
                    </div>
                </div>
            @endforeach
        </div>
        <button type="button" class="btn btn-secondary" id="add-question">Soru Ekle</button>
        <hr>
        <button type="submit" class="btn btn-success">Güncelle</button>
    </form>
</div>

<script>
    document.addEventListener('DOMContentLoaded', function() {
        let questionIndex = @json(count($exam->questions));

        document.getElementById('add-question').addEventListener('click', function() {
            questionIndex++;
            const container = document.getElementById('questions-container');
            const questionHTML = `
                <div class="question">
                    <h4>Soru ${questionIndex}</h4>
                    <div class="form-group">
                        <label for="questions[${questionIndex}][question_text]">Soru Metni</label>
                        <textarea class="form-control" id="questions[${questionIndex}][question_text]" name="questions[${questionIndex}][question_text]" required></textarea>
                    </div>
                    <div class="form-group">
                        <label for="questions[${questionIndex}][max_score]">Maksimum Puan</label>
                        <input type="number" class="form-control" id="questions[${questionIndex}][max_score]" name="questions[${questionIndex}][max_score]" required>
                    </div>
                    <div class="form-group">
                        <label>Seçenekler (Doğru cevabı seçiniz)</label>
                        <div class="options-container">
                            <div class="option">
                                <input type="radio" id="question_${questionIndex}_option_A" name="questions[${questionIndex}][correct_answer]" value="A" required>
                                <label for="question_${questionIndex}_option_A">Seçenek A</label>
                                <input type="text" class="form-control" name="questions[${questionIndex}][options][]" placeholder="Seçenek A" required>
                            </div>
                            <div class="option">
                                <input type="radio" id="question_${questionIndex}_option_B" name="questions[${questionIndex}][correct_answer]" value="B" required>
                                <label for="question_${questionIndex}_option_B">Seçenek B</label>
                                <input type="text" class="form-control" name="questions[${questionIndex}][options][]" placeholder="Seçenek B" required>
                            </div>
                            <div class="option">
                                <input type="radio" id="question_${questionIndex}_option_C" name="questions[${questionIndex}][correct_answer]" value="C" required>
                                <label for="question_${questionIndex}_option_C">Seçenek C</label>
                                <input type="text" class="form-control" name="questions[${questionIndex}][options][]" placeholder="Seçenek C" required>
                            </div>
                        </div>
                        <button type="button" class="btn btn-secondary add-option" data-question-index="${questionIndex}">Seçenek Ekle</button>
                    </div>
                </div>
            `;
            container.insertAdjacentHTML('beforeend', questionHTML);
        });

        // Seçenek Ekleme
        document.addEventListener('click', function(event) {
            if (event.target.classList.contains('add-option')) {
                const questionIndex = event.target.getAttribute('data-question-index');
                const optionCount = document.querySelectorAll(`.question:nth-child(${parseInt(questionIndex) + 1}) .options-container .option`).length;
                const newOptionIndex = String.fromCharCode(65 + optionCount); // 65 = 'A' ASCII değeri
                const optionHTML = `
                    <div class="option">
                        <input type="radio" id="question_${questionIndex}_option_${newOptionIndex}" name="questions[${questionIndex}][correct_answer]" value="${newOptionIndex}" required>
                        <label for="question_${questionIndex}_option_${newOptionIndex}">Seçenek ${newOptionIndex}</label>
                        <input type="text" class="form-control" name="questions[${questionIndex}][options][]" placeholder="Seçenek ${newOptionIndex}" required>
                    </div>
                `;
                event.target.parentNode.querySelector('.options-container').insertAdjacentHTML('beforeend', optionHTML);
            }
        });
    });
</script>
@endsection
