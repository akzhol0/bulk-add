const test = require('node:test');
const assert = require('node:assert/strict');

const {
  normalizeCourseNames,
  parseArgs,
  resolveCourseNames,
} = require('./add-course');

test('reads multiple unique course names from config', () => {
  assert.deepEqual(
    resolveCourseNames(null, {
      courseNames: [' Course A ', 'Course B', 'Course A', ''],
      courseName: 'Old course',
    }),
    ['Course A', 'Course B'],
  );
});

test('keeps backward compatibility with a single courseName', () => {
  assert.deepEqual(resolveCourseNames(null, { courseName: 'Old course' }), ['Old course']);
});

test('command line course overrides the config array', () => {
  assert.deepEqual(
    resolveCourseNames('CLI course', { courseNames: ['Course A', 'Course B'] }),
    ['CLI course'],
  );
});

test('normalizes empty and duplicate course names', () => {
  assert.deepEqual(normalizeCourseNames(['A', ' A ', null, '', 'B']), ['A', 'B']);
});

test('still parses the existing add command options', () => {
  assert.deepEqual(
    parseArgs(['--slug=program-1', '--dry-run', '--course=Course A', '--limit=2']),
    {
      dryRun: true,
      slugs: ['program-1'],
      limit: 2,
      course: 'Course A',
    },
  );
});
