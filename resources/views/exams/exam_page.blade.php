<!DOCTYPE html>
<html lang="tr">

<head>
    <!-- Meta ve Stil Tanımları -->
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ $exam->title }} - Sınav Sayfası</title>
    <style>
        /* CSS Stilleri */
        body,
        html {
            height: 100%;
            margin: 0;
            font-family: Arial, sans-serif;
        }

        .video-background {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            overflow: hidden;
            z-index: -1;
        }

        video {
            width: 100%;
            height: 100%;
            object-fit: cover;
        }

        .exam-title {
            text-align: center;
            font-size: 32px;
            font-weight: bold;
            color: #2c3e50;
            margin-top: 20px;
        }

        .warning-wrapper {
            background-color: rgba(255, 255, 255, 0.9);
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 0 20px rgba(0, 0, 0, 0.2);
            width: 80%;
            max-width: 500px;
            text-align: center;
            margin: auto;
            position: relative;
            top: 50%;
            transform: translateY(-50%);
        }

        #timer {
            position: absolute;
            top: 10px;
            right: 10px;
            background: rgba(0, 0, 0, 0.7);
            color: white;
            padding: 10px;
            border-radius: 5px;
        }

        button {
            font-size: 16px;
            font-weight: bold;
            background-color: #28a745;
            color: white;
            border: none;
            border-radius: 5px;
            padding: 12px 25px;
            cursor: pointer;
            transition: background-color 0.3s;
            margin-top: 20px;
        }

        button:hover {
            background-color: #218838;
        }

        .btn-danger {
            background-color: #dc3545;
        }

        .btn-danger:hover {
            background-color: #c82333;
        }

        .exam-container {
            display: none;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            background-color: white;
            padding: 20px;
        }

        .question-container {
            display: none;
            margin-bottom: 20px;
            text-align: center;
        }

        .question-container.active {
            display: block;
        }

        .question-image {
            text-align: center;
            margin-bottom: 10px;
        }

        .question-image img {
            max-width: 100%;
            max-height: 300px;
        }

        .question-text {
            font-weight: bold;
            margin-bottom: 20px;
            font-size: 18px;
            text-align: center;
        }

        .option-label {
            font-size: 16px;
            margin-bottom: 10px;
        }

        .has-image .options-container {
            display: flex;
            flex-direction: column;
            align-items: flex-start;
            margin-top: 10px;
        }

        .has-image .option-label {
            text-align: left;
        }
    </style>
</head>

<body>
    <!-- Video Arka Planı -->
    <div class="video-background">
        <video autoplay muted loop>
            <source src="{{ asset('assets/videos/background.mp4') }}" type="video/mp4">
            Tarayıcınız video etiketini desteklemiyor.
        </video>
    </div>

    <!-- Zamanlayıcı -->
    <div id="timer">00:00</div>

    <!-- Uyarı ve Sınava Başlama Butonu -->
    <div id="warningContainer" class="warning-wrapper">
        <h1>Sınavınızda Başarılar</h1>
        <button id="startExam" class="btn btn-primary">Sınava Başla</button>

        <!-- Admin için "Sınavı Bitir" Butonu -->
        @if(auth()->user()->hasRole('admin'))
        <button id="endExam" class="btn btn-danger">Sınavı Bitir</button>
        @endif
    </div>

    <!-- Sınav İçeriği -->
    <div class="exam-container" id="examContainer">
        <div id="examContent">
            <!-- Sınav Başlığı Ortalanmış Şekilde -->
            <h1 class="exam-title">{{ $exam->title }}</h1>

            <form id="examForm" action="{{ route('exam.submit', $exam->id) }}" method="POST">
                @csrf
                <div id="questionsContainer" class="question-wrapper">
                    @foreach($questions as $index => $question)
                    @php
                    $options = is_array($question->options) ? $question->options : json_decode($question->options, true);
                    $optionLetters = ['A', 'B', 'C', 'D', 'E']; // Şıklar için harfler
                    @endphp

                    <!-- Soruya ait görsel varsa özel stil sınıfı ekleniyor -->
                    <div class="question-container {{ $index == 0 ? 'active' : '' }} {{ $question->image ? 'has-image' : '' }}"
                        id="question-{{ $index }}">
                        <!-- Soruya ait görsel varsa burada gösterilecek -->
                        @if ($question->image)
                        <div class="question-image">
                            <img src="{{ Storage::url($question->image) }}" alt="Soru Görseli">
                        </div>
                        @endif

                        <p class="question-text">{{ $question->question_text }}</p>
                        <input type="hidden" name="questions[{{ $question->id }}][points]"
                            value="{{ $question->max_score ?? 10 }}"> <!-- Soru puanı -->

                        <div class="options-container">
                            @foreach($options as $optionIndex => $option)
                            @php
                            $optionLetter = $optionLetters[$optionIndex] ?? 'Unknown';
                            @endphp
                            <div>
                                <input type="radio" id="option{{ $index }}-{{ $optionIndex }}"
                                    name="questions[{{ $question->id }}][answer]" value="{{ $optionLetter }}">
                                <label class="option-label" for="option{{ $index }}-{{ $optionIndex }}">
                                    {{ $optionLetter }}: {{ $option }}

                                    @if(isset($option['image']))
                                    <div class="option-image">
                                        <img src="{{ Storage::url('exam_images/' . $option['image']) }}"
                                            alt="Şık Görseli">
                                    </div>
                                    @endif

                                </label>
                            </div>
                            @endforeach
                        </div>

                        @if($index < count($questions) - 1)
                        <button type="button" class="btn btn-primary next-question" data-next="{{ $index + 1 }}">Sonraki
                            Soru</button>
                        @else
                        <button type="button" id="finishExam" class="btn btn-success">Sınavı Bitir</button>
                        @endif
                    </div>
                    @endforeach
                </div>
            </form>
        </div>
    </div>

    <!-- JavaScript Kodları -->
    <script>
        document.addEventListener('DOMContentLoaded', function () {
            function enterFullscreen() {
                var elem = document.documentElement;
                if (elem.requestFullscreen) {
                    elem.requestFullscreen();
                } else if (elem.mozRequestFullScreen) { // Firefox
                    elem.mozRequestFullScreen();
                } else if (elem.webkitRequestFullscreen) { // Chrome, Safari and Opera
                    elem.webkitRequestFullscreen();
                } else if (elem.msRequestFullscreen) { // IE/Edge
                    elem.msRequestFullscreen();
                }
            }
    
            function exitFullscreen() {
                if (document.exitFullscreen) {
                    document.exitFullscreen();
                } else if (document.mozCancelFullScreen) { // Firefox
                    document.mozCancelFullScreen();
                } else if (document.webkitExitFullscreen) { // Chrome, Safari and Opera
                    document.webkitExitFullscreen();
                } else if (document.msExitFullscreen) { // IE/Edge
                    document.msExitFullscreen();
                }
            }
    
            const startExamButton = document.getElementById('startExam');
            const warningContainer = document.getElementById('warningContainer');
            const examContainer = document.getElementById('examContainer');
            const endExamButton = document.getElementById('endExam'); // Sınavı Bitir butonu
    
            // Veritabanından alınan sınav süresi
            const examDuration = {{ $exam->duration * 60 }}; // Dakikayı saniyeye çevir
    
            startExamButton.addEventListener('click', function () {
                enterFullscreen();
                warningContainer.style.display = 'none';
                examContainer.style.display = 'flex';
                startTimer(examDuration); // Veritabanından alınan süre ile başlat
            });
    
            if (endExamButton) {
                // Admin butona tıkladığında sınavı sonlandır
                endExamButton.addEventListener('click', function() {
                    document.getElementById('examForm').submit();
                    exitFullscreen();
                });
            }
    
            let currentQuestion = 0;
            const totalQuestions = {{ count($questions) }};
            const nextQuestionButtons = document.querySelectorAll('.next-question');
            const finishButton = document.getElementById('finishExam');
    
            nextQuestionButtons.forEach(button => {
                button.addEventListener('click', function () {
                    const currentQuestionInputs = document.querySelectorAll(`#question-${currentQuestion} input[type="radio"]`);
                    let isAnswered = false;
    
                    currentQuestionInputs.forEach(input => {
                        if (input.checked) {
                            isAnswered = true;
                        }
                    });
    
                    // Eğer cevap işaretlenmemişse uyarı göster
                    if (!isAnswered) {
                        const warningMessage = document.createElement('div');
                        warningMessage.textContent = 'Lütfen bir şık işaretleyin.';
                        warningMessage.style.position = 'fixed';
                        warningMessage.style.top = '50%';
                        warningMessage.style.left = '50%';
                        warningMessage.style.transform = 'translate(-50%, -50%)';
                        warningMessage.style.padding = '20px';
                        warningMessage.style.backgroundColor = 'rgba(255, 0, 0, 0.8)';
                        warningMessage.style.color = 'white';
                        warningMessage.style.borderRadius = '5px';
                        warningMessage.style.zIndex = '9999';
    
                        document.body.appendChild(warningMessage);
    
                        // 2 saniye sonra uyarıyı kaldır
                        setTimeout(() => {
                            document.body.removeChild(warningMessage);
                        }, 2000);
    
                        return; // Bir şık işaretlenmediği için sonraki soruya geçilmez
                    }
    
                    // Eğer cevap işaretlenmişse bir sonraki soruya geç
                    document.querySelector(`#question-${currentQuestion}`).classList.remove('active');
                    currentQuestion++;
                    if (currentQuestion < totalQuestions) {
                        document.querySelector(`#question-${currentQuestion}`).classList.add('active');
                    } else {
                        finishButton.style.display = 'block';
                    }
                });
            });
    
            if (finishButton) {
                finishButton.addEventListener('click', function () {
                    const examForm = document.getElementById('examForm');
                    examForm.submit();
                    exitFullscreen();
                });
            }
    
            function startTimer(duration) {
                const timerDisplay = document.getElementById('timer');
                let timer = duration, minutes, seconds;
    
                const interval = setInterval(function () {
                    minutes = parseInt(timer / 60, 10);
                    seconds = parseInt(timer % 60, 10);
    
                    minutes = minutes < 10 ? "0" + minutes : minutes;
                    seconds = seconds < 10 ? "0" + seconds : seconds;
    
                    timerDisplay.textContent = minutes + ":" + seconds;
    
                    if (--timer < 0) {
                        timer = 0;
                        clearInterval(interval);
                        document.getElementById('examForm').submit(); // Süre bitince form submit edilir
                        exitFullscreen();
                    }
                }, 1000);
            }
        });
    </script>    
</body>
</html>


