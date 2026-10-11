// Track manual user inputs to prevent overwriting user-typed values
let userInputs = new Set();

const fields = [
    'start_date', 'end_date', 'today_date', 'camp_days', 'rem_days', 'net_budget', 
    'currency_code', 'cost', 'ex_rate', 'conv_cost',
    'imp', 'clicks', 'reach', 'freq', 'cpm', 'cpc', 'ctr',
    'leads', 'action', 'conv', 'cpl', 'cpa', 'cps', 'cvr', 'cr', 'upt',
    'sales', 'rev', 'ttv', 'visitors', 'atc', 'aov', 'car', 'gmv', 'trans', 'total_web', 'visits', 'cpviews', 'cpv', 'items', 'carts',
    'eng', 'cpe', 'er', 'grps', 'roas', 'roi'
];

const types = {
    money: ['net_budget', 'cost', 'conv_cost', 'cpm', 'cpc', 'cpl', 'cpa', 'cps', 'rev', 'cpv', 'ttv', 'gmv', 'aov', 'cpe'],
    percent: ['ctr', 'er', 'cvr', 'roi', 'car'],
    count: ['camp_days', 'rem_days', 'imp', 'clicks', 'leads', 'action', 'eng', 'sales', 'conv', 'reach', 'visitors', 'atc', 'trans', 'total_web', 'visits', 'cpviews', 'items', 'carts', 'cr'],
    decimal: ['ex_rate', 'freq', 'grps', 'roas', 'upt'],
    text: ['currency_code'],
    date: ['start_date', 'end_date', 'today_date']
};

function getVal(id) {
    if (types.text.includes(id) || types.date.includes(id)) {
        return document.getElementById(id).value || null;
    }
    const val = parseFloat(document.getElementById(id).value);
    return isNaN(val) ? null : val;
}

function setVal(id, val) {
    const el = document.getElementById(id);
    if (document.activeElement === el || userInputs.has(id)) return;
    
    if (types.text.includes(id) || types.date.includes(id)) {
        if (val === null) {
            if (el.value !== '') el.value = '';
            return;
        }
        if (el.value !== val) el.value = val;
        return;
    }

    if (val === null || !isFinite(val) || isNaN(val)) {
        if (el.value !== '') el.value = '';
        return;
    }

    let formattedVal;
    if (types.money.includes(id) || types.percent.includes(id)) {
        formattedVal = val.toFixed(2);
    } else if (types.count.includes(id)) {
        formattedVal = Math.round(val);
    } else if (id === 'ex_rate') {
        formattedVal = val.toFixed(4);
    } else {
        formattedVal = val.toFixed(2); 
    }

    if (parseFloat(el.value) !== parseFloat(formattedVal)) {
        el.value = formattedVal;
        el.classList.remove('calculated-flash');
        void el.offsetWidth; // Trigger reflow
        el.classList.add('calculated-flash');
    }
}

function runEngine() {
    let data = {};
    fields.forEach(f => data[f] = getVal(f));

    let changed = true;
    let iterations = 0;
    const MAX_ITERATIONS = 25;

    while (changed && iterations < MAX_ITERATIONS) {
        changed = false;

        const tryCalc = (key, calculatedValue) => {
            if (data[key] === null && !userInputs.has(key) && !isNaN(calculatedValue) && isFinite(calculatedValue)) {
                if (key === 'roi' || calculatedValue >= 0) { 
                    data[key] = calculatedValue;
                    changed = true;
                }
            }
        };
        
        const safeDiv = (a, b) => (b && b !== 0) ? a / b : null;

        // --- Metric Synchronizations ---
        if (data.gmv !== null) tryCalc('ttv', data.gmv);
        if (data.ttv !== null) tryCalc('gmv', data.ttv);

        // --- 1. Campaign Timeline & Budget Math ---
        let start = data.start_date ? new Date(data.start_date + 'T00:00:00') : null;
        let end = data.end_date ? new Date(data.end_date + 'T00:00:00') : null;
        let today = data.today_date ? new Date(data.today_date + 'T00:00:00') : null;
        
        if (start && end && end >= start) {
            let diffDays = Math.ceil(Math.abs(end - start) / (1000 * 60 * 60 * 24)) + 1;
            tryCalc('camp_days', diffDays);
        }
        
        if (end && today) {
            let rDays = Math.ceil((end - today) / (1000 * 60 * 60 * 24));
            tryCalc('rem_days', Math.max(0, rDays));
        }

        // --- 2. Cost & Currency Math ---
        if (data.currency_code !== null) {
            document.getElementById('curr_prefix').innerText = data.currency_code;
        }
        if (data.cost !== null && data.ex_rate !== null) tryCalc('conv_cost', data.cost * data.ex_rate);
        if (data.conv_cost !== null && data.ex_rate !== null) tryCalc('cost', safeDiv(data.conv_cost, data.ex_rate));
        if (data.conv_cost !== null && data.cost !== null) tryCalc('ex_rate', safeDiv(data.conv_cost, data.cost));

        // --- 3. Media Delivery ---
        if (data.cost !== null && data.imp !== null) tryCalc('cpm', safeDiv(data.cost, (data.imp / 1000)));
        if (data.cost !== null && data.cpm !== null) tryCalc('imp', safeDiv(data.cost, data.cpm) * 1000);
        if (data.cpm !== null && data.imp !== null) tryCalc('cost', (data.cpm * data.imp) / 1000);

        if (data.cost !== null && data.clicks !== null) tryCalc('cpc', safeDiv(data.cost, data.clicks));
        if (data.cost !== null && data.cpc !== null) tryCalc('clicks', safeDiv(data.cost, data.cpc));
        if (data.cpc !== null && data.clicks !== null) tryCalc('cost', data.cpc * data.clicks);

        if (data.clicks !== null && data.imp !== null) tryCalc('ctr', safeDiv(data.clicks, data.imp) * 100);
        if (data.imp !== null && data.ctr !== null) tryCalc('clicks', data.imp * (data.ctr / 100));
        if (data.clicks !== null && data.ctr !== null) tryCalc('imp', safeDiv(data.clicks, (data.ctr / 100)));

        if (data.imp !== null && data.reach !== null) tryCalc('freq', safeDiv(data.imp, data.reach));
        if (data.reach !== null && data.freq !== null) tryCalc('imp', data.reach * data.freq);
        if (data.imp !== null && data.freq !== null) tryCalc('reach', safeDiv(data.imp, data.freq));

        // --- 4. Action & Conversions ---
        if (data.cost !== null && data.leads !== null) tryCalc('cpl', safeDiv(data.cost, data.leads));
        if (data.cost !== null && data.cpl !== null) tryCalc('leads', safeDiv(data.cost, data.cpl));
        if (data.cpl !== null && data.leads !== null) tryCalc('cost', data.cpl * data.leads);

        if (data.cost !== null && data.action !== null) tryCalc('cpa', safeDiv(data.cost, data.action));
        if (data.cost !== null && data.cpa !== null) tryCalc('action', safeDiv(data.cost, data.cpa));
        if (data.cpa !== null && data.action !== null) tryCalc('cost', data.cpa * data.action);

        if (data.cost !== null && data.sales !== null) tryCalc('cps', safeDiv(data.cost, data.sales));
        if (data.cost !== null && data.cps !== null) tryCalc('sales', safeDiv(data.cost, data.cps));
        if (data.cps !== null && data.sales !== null) tryCalc('cost', data.cps * data.sales);

        if (data.conv !== null && data.clicks !== null) tryCalc('cvr', safeDiv(data.conv, data.clicks) * 100);
        if (data.clicks !== null && data.cvr !== null) tryCalc('conv', data.clicks * (data.cvr / 100));

        if (data.items !== null && data.trans !== null) tryCalc('upt', safeDiv(data.items, data.trans));
        if (data.upt !== null && data.trans !== null) tryCalc('items', data.upt * data.trans);
        if (data.items !== null && data.upt !== null) tryCalc('trans', safeDiv(data.items, data.upt));

        // --- 5. E-commerce & Web ---
        if (data.rev !== null && data.trans !== null) tryCalc('aov', safeDiv(data.rev, data.trans));
        if (data.rev !== null && data.aov !== null) tryCalc('trans', safeDiv(data.rev, data.aov));
        if (data.aov !== null && data.trans !== null) tryCalc('rev', data.aov * data.trans);

        if (data.trans !== null && data.aov !== null) tryCalc('gmv', data.trans * data.aov);
        if (data.gmv !== null && data.trans !== null) tryCalc('aov', safeDiv(data.gmv, data.trans));
        if (data.gmv !== null && data.aov !== null) tryCalc('trans', safeDiv(data.gmv, data.aov));

        if (data.carts !== null && data.trans !== null) {
            tryCalc('car', safeDiv((data.carts - data.trans), data.carts) * 100);
        }
        if (data.carts !== null && data.car !== null) {
            tryCalc('trans', data.carts - (data.carts * (data.car / 100)));
        }
        if (data.trans !== null && data.car !== null) {
            tryCalc('carts', safeDiv(data.trans, (1 - (data.car / 100))));
        }

        if (data.cost !== null && data.cpviews !== null) tryCalc('cpv', safeDiv(data.cost, data.cpviews));
        if (data.cost !== null && data.cpv !== null) tryCalc('cpviews', safeDiv(data.cost, data.cpv));
        if (data.cpv !== null && data.cpviews !== null) tryCalc('cost', data.cpv * data.cpviews);

        // --- 6. Engagement & Returns ---
        if (data.cost !== null && data.eng !== null) tryCalc('cpe', safeDiv(data.cost, data.eng));
        if (data.cost !== null && data.cpe !== null) tryCalc('eng', safeDiv(data.cost, data.cpe));
        if (data.cpe !== null && data.eng !== null) tryCalc('cost', data.cpe * data.eng);

        if (data.eng !== null && data.imp !== null) tryCalc('er', safeDiv(data.eng, data.imp) * 100);
        if (data.imp !== null && data.er !== null) tryCalc('eng', data.imp * (data.er / 100));

        if (data.rev !== null && data.cost !== null) tryCalc('roas', safeDiv(data.rev, data.cost));
        if (data.cost !== null && data.roas !== null) tryCalc('rev', data.cost * data.roas);
        if (data.rev !== null && data.roas !== null) tryCalc('cost', safeDiv(data.rev, data.roas));

        if (data.rev !== null && data.cost !== null) tryCalc('roi', safeDiv((data.rev - data.cost), data.cost) * 100);
        if (data.cost !== null && data.roi !== null) tryCalc('rev', data.cost * ((data.roi / 100) + 1));

        iterations++;
    }

    // Write back calculated values
    fields.forEach(f => {
        if (!userInputs.has(f)) {
            setVal(f, data[f]);
        }
    });
}

// Attach Event Listeners
document.querySelectorAll('.calc-input').forEach(input => {
    input.addEventListener('input', (e) => {
        const id = e.target.id;
        if (e.target.value.trim() === '') {
            userInputs.delete(id);
        } else {
            userInputs.add(id);
        }
        runEngine();
    });
    
    input.addEventListener('blur', (e) => {
        const id = e.target.id;
        const val = getVal(id);
        if (val !== null && !types.date.includes(id)) {
            userInputs.delete(id); 
            setVal(id, val); 
            userInputs.add(id);
        }
    });
});

// Initialize Defaults on Load
window.onload = () => {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    
    const todayString = `${year}-${month}-${day}`;
    document.getElementById('today_date').value = todayString;
    
    document.getElementById('curr_prefix').innerText = document.getElementById('currency_code').value;
    
    userInputs.add('today_date');
    runEngine();
};

// Reset All Function
function clearAll() {
    userInputs.clear();
    document.querySelectorAll('.calc-input').forEach(input => {
        if (input.id !== 'currency_code' && input.id !== 'today_date') {
            input.value = '';
        }
        input.classList.remove('calculated-flash');
    });
    
    document.getElementById('currency_code').value = 'USD';
    document.getElementById('curr_prefix').innerText = 'USD';
    
    const now = new Date();
    const todayString = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
    document.getElementById('today_date').value = todayString;
    
    userInputs.add('today_date');
    runEngine();
}