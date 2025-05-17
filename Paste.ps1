$decodedContent = $decodedContent -replace '<script[^>]*language\s*=\s*["'']?javascript["'']?[^>]*>.*?<\/script>', '', 'Singleline,IgnoreCase'
