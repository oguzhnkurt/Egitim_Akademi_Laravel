<!DOCTYPE html>
<html lang="tr">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ $exam->title }} - Sınav</title>
    <style>
        body,
        html {
            height: 100%;
            margin: 0;
            display: flex;
            justify-content: center;
            align-items: center;
            background-color: #fff;
        }

        .exam-container {
            width: 100%;
            max-width: 600px;
            padding: 20px;
            border: 1px solid #ccc;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            background-color: #f9f9f9;
        }

        .question-container {
            display: none;
        }

        .question-container.active {
            display: block;
        }
    </style>
</head>

<body>
    <div class="exam-container">
        <h1>{{ $exam->title }}</h1>
        <form id="examForm" action="{{ route('exam.submit', $exam->id) }}" method="POST">
            @csrf
            <div id="examContainer">
                @foreach($questions as $index => $question)
                <div class="question-container" id="question-{{ $index }}"
                    style="display: {{ $index == 0 ? 'block' : 'none' }}">
                    <p>{{ $question->question_text }}</p>
                    <input type="hidden" name="questions[{{ $question->id }}][id]" value="{{ $question->id }}">
                    @foreach($question->options as $option)
                    <div>
                        <input type="radio" id="option{{ $option->id }}" name="questions[{{ $question->id }}][answer]"
                            value="{{ $option->option_text }}">
                        <label for="option{{ $option->id }}">
                            {{ $option->option_text }}
                            @if($option->image)
                            <br>
                            <img src="{{ asset('storage/' . $option->image) }}" alt="Şık Görseli"
                                style="max-width: 100%; height: auto;">
                            @endif
                        </label>
                    </div>
                    @endforeach

                    <button type="button" class="btn btn-primary next-question" data-next="{{ $index + 1 }}">Sonraki
                        Soru</button>
                </div>
                @endforeach
            </div>
            <button type="submit" style="display: none;" id="submitExam">Sınavı Tamamla</button>
        </form>
    </div>

    <script>
        document.addEventListener('DOMContentLoaded', function () {
            const nextButtons = document.querySelectorAll('.next-question');
            let currentQuestionIndex = 0;

            nextButtons.forEach(button => {
                button.addEventListener('click', function () {
                    const nextIndex = parseInt(this.getAttribute('data-next'));
                    const currentContainer = document.getElementById('question-' + currentQuestionIndex);
                    const nextContainer = document.getElementById('question-' + nextIndex);

                    if (nextContainer) {
                        currentContainer.style.display = 'none';
                        nextContainer.style.display = 'block';
                        currentQuestionIndex = nextIndex;
                    } else {
                        document.getElementById('submitExam').click();
                    }
                });
            });

            // Tam ekran modunu başlat
            function enterFullScreen() {
                if (document.documentElement.requestFullscreen) {
                    document.documentElement.requestFullscreen();
                } else if (document.documentElement.mozRequestFullScreen) {
                    document.documentElement.mozRequestFullScreen();
                } else if (document.documentElement.webkitRequestFullscreen) {
                    document.documentElement.webkitRequestFullscreen();
                } else if (document.documentElement.msRequestFullscreen) {
                    document.documentElement.msRequestFullscreen();
                }
            }

            // Tam ekranı kullanıcı etkileşimi ile başlat
            document.querySelector('.next-question').addEventListener('click', enterFullScreen);
            enterFullScreen(); // Sayfa yüklendiğinde de tam ekran moduna geçiş yap
        });
    </script>
</body>

</html>