import dayjs from 'dayjs';
import weekOfYear from 'dayjs/plugin/weekOfYear.js';
import isoWeek from 'dayjs/plugin/isoWeek.js';

dayjs.extend(weekOfYear);
dayjs.extend(isoWeek);

const dates = [
    '2025-12-31',
    '2025-12-30',
    '2025-01-01',
    '2025-12-28'
];

console.log('Using weekOfYear (locale default):');
dates.forEach(d => {
    console.log(`${d}: ${dayjs(d).week()}`);
});

console.log('\nUsing isoWeek:');
dates.forEach(d => {
    console.log(`${d}: ${dayjs(d).isoWeek()}`);
});
