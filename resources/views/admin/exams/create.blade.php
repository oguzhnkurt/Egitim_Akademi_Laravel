@extends('layouts.master')

@section('content')

<div class="container">
    <h1>Yeni Sınav Oluştur</h1>
    <form action="{{ route('admin.exams.store') }}" method="POST" enctype="multipart/form-data">
        @csrf
        <div class="form-group">
            <label for="title">Başlık</label>
            <input type="text" class="form-control" id="title" name="title" required>
        </div>
        <div class="form-group">
            <label for="description">Açıklama</label>
            <textarea class="form-control" id="description" name="description" required></textarea>
        </div>
        <div class="form-group">
            <label for="start_date">Başlangıç Tarihi</label>
            <input type="date" class="form-control" id="start_date" name="start_date" required>
        </div>
        <div class="form-group">
            <label for="end_date">Bitiş Tarihi</label>
            <input type="date" class="form-control" id="end_date" name="end_date" required>
        </div>
        <div class="form-group">
            <label for="duration">Süre (dakika)</label>
            <input type="number" class="form-control" id="duration" name="duration" required>
        </div>

        <hr>
        <h3>Sorular</h3>
        <div id="questions-container">
            <div class="question" data-index="0">
                <h4>Soru 1</h4>
                <div class="form-group">
                    <label for="questions[0][question_text]">Soru Metni</label>
                    <textarea class="form-control" id="questions[0][question_text]" name="questions[0][question_text]" required></textarea>
                </div>
                <div class="form-group">
                    <label for="questions[0][max_score]">Maksimum Puan</label>
                    <input type="number" class="form-control" id="questions[0][max_score]" name="questions[0][max_score]" required>
                </div>

                <!-- Görsel Yükleme Alanı -->
                <div class="form-group">
                    <label for="questions[0][image]">Görsel Ekle</label>
                    <input type="file" class="form-control image-upload" id="questions[0][image]" name="questions[0][image]" data-question-index="0" accept="image/*">
                    <div id="uploaded-image-0"></div> <!-- Görsel önizlemesi -->
                </div>

                <!-- Seçenek Görsel Ekleme Tercihi -->
                <div class="form-group">
                    <input type="checkbox" id="add-option-images-0" class="add-option-images-checkbox" data-question-index="0">
                    <label for="add-option-images-0">Şıklara görsel ekle</label>
                </div>

                <div class="form-group">
                    <label>Seçenekler (Doğru cevabı seçiniz)</label>
                    <div class="options-container">
                        <div class="option">
                            <input type="radio" id="question_0_option_A" name="questions[0][correct_option]" value="A" required>
                            <label for="question_0_option_A">Seçenek A</label>
                            <input type="text" class="form-control" name="questions[0][options][]" placeholder="Seçenek A" required>
                            <input type="file" class="form-control option-image-upload d-none" id="question_0_option_A_image" name="questions[0][option_images][]" accept="image/*">
                        </div>
                        <div class="option">
                            <input type="radio" id="question_0_option_B" name="questions[0][correct_option]" value="B">
                            <label for="question_0_option_B">Seçenek B</label>
                            <input type="text" class="form-control" name="questions[0][options][]" placeholder="Seçenek B" required>
                            <input type="file" class="form-control option-image-upload d-none" id="question_0_option_B_image" name="questions[0][option_images][]" accept="image/*">
                        </div>
                        <div class="option">
                            <input type="radio" id="question_0_option_C" name="questions[0][correct_option]" value="C">
                            <label for="question_0_option_C">Seçenek C</label>
                            <input type="text" class="form-control" name="questions[0][options][]" placeholder="Seçenek C" required>
                            <input type="file" class="form-control option-image-upload d-none" id="question_0_option_C_image" name="questions[0][option_images][]" accept="image/*">
                        </div>
                        <div class="option">
                            <input type="radio" id="question_0_option_D" name="questions[0][correct_option]" value="D">
                            <label for="question_0_option_D">Seçenek D</label>
                            <input type="text" class="form-control" name="questions[0][options][]" placeholder="Seçenek D" required>
                            <input type="file" class="form-control option-image-upload d-none" id="question_0_option_D_image" name="questions[0][option_images][]" accept="image/*">
                        </div>
                    </div>
                    <button type="button" class="btn btn-secondary add-option" data-question-index="0">Seçenek Ekle</button>
                </div>
            </div>
        </div>
        <button type="button" class="btn btn-secondary" id="add-question">Soru Ekle</button>
        <hr>
        <button type="submit" class="btn btn-primary">Sınavı Kaydet</button>
    </form>
</div>

<script>
document.addEventListener('DOMContentLoaded', function() {
    let questionIndex = 1;

    document.getElementById('add-question').addEventListener('click', function() {
        const container = document.getElementById('questions-container');
        const questionHTML = `
            <div class="question" data-index="${questionIndex}">
                <h4>Soru ${questionIndex + 1}</h4>
                <div class="form-group">
                    <label for="questions[${questionIndex}][question_text]">Soru Metni</label>
                    <textarea class="form-control" id="questions[${questionIndex}][question_text]" name="questions[${questionIndex}][question_text]" required></textarea>
                </div>
                <div class="form-group">
                    <label for="questions[${questionIndex}][max_score]">Maksimum Puan</label>
                    <input type="number" class="form-control" id="questions[${questionIndex}][max_score]" name="questions[${questionIndex}][max_score]" required>
                </div>

                <!-- Görsel Yükleme Alanı -->
                <div class="form-group">
                    <label for="questions[${questionIndex}][image]">Görsel Ekle</label>
                    <input type="file" class="form-control image-upload" id="questions[${questionIndex}][image]" name="questions[${questionIndex}][image]" data-question-index="${questionIndex}" accept="image/*">
                    <div id="uploaded-image-${questionIndex}"></div> <!-- Görsel önizlemesi -->
                </div>

                <!-- Seçenek Görsel Ekleme Tercihi -->
                <div class="form-group">
                    <input type="checkbox" id="add-option-images-${questionIndex}" class="add-option-images-checkbox" data-question-index="${questionIndex}">
                    <label for="add-option-images-${questionIndex}">Şıklara görsel ekle</label>
                </div>

                <div class="form-group">
                    <label>Seçenekler (Doğru cevabı seçiniz)</label>
                    <div class="options-container">
                        <div class="option">
                            <input type="radio" id="question_${questionIndex}_option_A" name="questions[${questionIndex}][correct_option]" value="A" required>
                            <label for="question_${questionIndex}_option_A">Seçenek A</label>
                            <input type="text" class="form-control" name="questions[${questionIndex}][options][]" placeholder="Seçenek A" required>
                            <input type="file" class="form-control option-image-upload d-none" id="question_${questionIndex}_option_A_image" name="questions[${questionIndex}][option_images][]" accept="image/*">
                        </div>
                        <div class="option">
                            <input type="radio" id="question_${questionIndex}_option_B" name="questions[${questionIndex}][correct_option]" value="B">
                            <label for="question_${questionIndex}_option_B">Seçenek B</label>
                            <input type="text" class="form-control" name="questions[${questionIndex}][options][]" placeholder="Seçenek B" required>
                            <input type="file" class="form-control option-image-upload d-none" id="question_${questionIndex}_option_B_image" name="questions[${questionIndex}][option_images][]" accept="image/*">
                        </div>
                        <div class="option">
                            <input type="radio" id="question_${questionIndex}_option_C" name="questions[${questionIndex}][correct_option]" value="C">
                            <label for="question_${questionIndex}_option_C">Seçenek C</label>
                            <input type="text" class="form-control" name="questions[${questionIndex}][options][]" placeholder="Seçenek C" required>
                            <input type="file" class="form-control option-image-upload d-none" id="question_${questionIndex}_option_C_image" name="questions[${questionIndex}][option_images][]" accept="image/*">
                        </div>
                        <div class="option">
                            <input type="radio" id="question_${questionIndex}_option_D" name="questions[${questionIndex}][correct_option]" value="D">
                            <label for="question_${questionIndex}_option_D">Seçenek D</label>
                            <input type="text" class="form-control" name="questions[${questionIndex}][options][]" placeholder="Seçenek D" required>
                            <input type="file" class="form-control option-image-upload d-none" id="question_${questionIndex}_option_D_image" name="questions[${questionIndex}][option_images][]" accept="image/*">
                        </div>
                    </div>
                    <button type="button" class="btn btn-secondary add-option" data-question-index="${questionIndex}">Seçenek Ekle</button>
                </div>
            </div>
        `;
        container.insertAdjacentHTML('beforeend', questionHTML);
        questionIndex++;
    });

    // Görsel yükleme fonksiyonu
    document.addEventListener('change', function(event) {
        if (event.target.classList.contains('image-upload')) {
            const questionIndex = event.target.getAttribute('data-question-index');
            const formData = new FormData();
            formData.append('image', event.target.files[0]);

            axios.post('{{ route('admin.exams.uploadImage') }}', formData, {
                headers: {
                    'Content-Type': 'multipart/form-data'
                }
            }).then(function (response) {
                const imagePreview = document.getElementById(`uploaded-image-${questionIndex}`);
                imagePreview.innerHTML = `<img src="${response.data.path}" alt="Uploaded Image" style="max-width: 100%;">`;
            }).catch(function (error) {
                console.error('Görsel yüklenemedi:', error.response.data);
            });
        }
    });

    // Şıklar için görsel ekleme seçeneği kontrolü
    document.addEventListener('change', function(event) {
        if (event.target.classList.contains('add-option-images-checkbox')) {
            const questionIndex = event.target.getAttribute('data-question-index');
            const optionImagesInputs = document.querySelectorAll(`#questions-container [data-question-index="${questionIndex}"] .option-image-upload`);

            if (event.target.checked) {
                optionImagesInputs.forEach(input => input.classList.remove('d-none'));
            } else {
                optionImagesInputs.forEach(input => input.classList.add('d-none'));
            }
        }
    });

    // Şık ekleme fonksiyonu
    document.addEventListener('click', function(event) {
        if (event.target.classList.contains('add-option')) {
            const questionIndex = event.target.getAttribute('data-question-index');
            const optionsContainer = event.target.closest('.form-group').querySelector('.options-container');
            const optionCount = optionsContainer.querySelectorAll('.option').length;
            const newOptionIndex = String.fromCharCode(65 + optionCount); // 65 = 'A' ASCII değeri

            if (optionCount < 5) { // Bu kontrol ile 4'ten fazla seçeneğin eklenmesini engelleriz
                const optionHTML = `
                    <div class="option">
                        <input type="radio" id="question_${questionIndex}_option_${newOptionIndex}" name="questions[${questionIndex}][correct_option]" value="${newOptionIndex}" required>
                        <label for="question_${questionIndex}_option_${newOptionIndex}">Seçenek ${newOptionIndex}</label>
                        <input type="text" class="form-control" name="questions[${questionIndex}][options][]" placeholder="Seçenek ${newOptionIndex}" required>
                        <input type="file" class="form-control option-image-upload d-none" id="question_${questionIndex}_option_${newOptionIndex}_image" name="questions[${questionIndex}][option_images][]" data-option-index="${newOptionIndex}" accept="image/*">
                    </div>
                `;
                optionsContainer.insertAdjacentHTML('beforeend', optionHTML);
            }
        }
    });

});
</script>
@endsection


