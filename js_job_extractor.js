
window.__extractJobs = function() {
    const jobWords = ['entwickler', 'developer', 'administrator', 'engineer',
        'berater', 'consultant', 'trainee', 'junior', 'senior', 'azubi',
        'praktikant', 'werkstudent', 'ingenieur', 'fachinformatiker',
        'informatiker', 'programmierer', 'sachbearbeiter', 'assistent',
        'techniker', 'disponent', 'referent', 'controller', 'experte',
        'spezialist', 'specialist', 'scrum master', 'product owner', 'devops',
        'm/w/d', 'w/m/d', 'm/f/d', 'f/m/d', 'm/w', 'w/m', 'architekt',
        'architect', 'analyst', 'operator', 'manager', 'head', 'director',
        'koordinator', 'coordinator', 'projektmanager', 'project manager',
        'kaufmann', 'kauffrau', 'laborant', 'forscher', 'researcher'];
    
    function looksLikeJob(text) {
        const tl = text.toLowerCase();
        return jobWords.some(kw => tl.includes(kw));
    }
    
    const els = document.querySelectorAll('a, h2, h3, h4, h5, li, article, tr, .card, .teaser');
    const results = [];
    const seen = new Set();
    for (const el of els) {
        const text = (el.textContent || '').trim();
        if (text.length < 10 || text.length > 1500) continue;
        const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 3);
        const title = (lines[0] || text.substring(0, 150)).substring(0, 150);
        if (!looksLikeJob(title)) continue;
        const key = title.toLowerCase().substring(0, 50);
        if (seen.has(key)) continue;
        seen.add(key);
        const link = el.closest('a') || el.querySelector('a');
        const href = link ? link.href : (el.tagName === 'A' ? el.href : '');
        results.push({title: title, href: href || '', text: text.substring(0, 500)});
    }
    return results;
};

window.__extractLinks = function(keywords) {
    const results = [];
    const seen = new Set();
    document.querySelectorAll('a[href]').forEach(a => {
        const text = (a.textContent || '').trim().toLowerCase();
        const href = (a.getAttribute('href') || '').toLowerCase();
        const isMatch = keywords.some(kw => text.includes(kw) || href.includes(kw));
        if (isMatch && a.href && !seen.has(a.href)) {
            seen.add(a.href);
            results.push({href: a.href, text: (a.textContent || '').trim().substring(0, 100)});
        }
    });
    return results;
};
