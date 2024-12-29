<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Models\File;
use voku\helper\StopWords;
use voku\helper\ASCII;

class SearchController extends Controller
{
    public function search(Request $request)
    {
        $query = $request->input('query');

        // Stop words ve ASCII ile sorguyu sadeleştiriyoruz
        $simplifiedQuery = $this->simplifyQuery($query);

        // Veritabanında arama yapıyoruz
        $results = File::where('content', 'LIKE', "%{$simplifiedQuery}%")
            ->orWhere('title', 'LIKE', "%{$simplifiedQuery}%")
            ->get();

        return view('search.search-results', compact('results'));
    }

    // Sorguyu sadeleştir
    private function simplifyQuery($query)
    {
        // Küçük harfe dönüştürme
        $query = mb_strtolower($query);

        // StopWords sınıfının bir örneğini oluşturuyoruz
        $stopWords = new StopWords();

        // Türkçe stop word'leri alıyoruz
        $stopWordsList = $stopWords->getStopWordsFromLanguage('tr');

        // Sorguyu kelimelere ayır
        $words = explode(' ', $query);

        // Stop word'leri filtreliyoruz
        $filteredWords = array_diff($words, $stopWordsList);

        // ASCII karakter dönüşümü
        $simplifiedQuery = ASCII::to_ascii(implode(' ', $filteredWords));

        return $simplifiedQuery;
    }
}
