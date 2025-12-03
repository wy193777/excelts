/**
 * Unzipper parse-datetime module
 * Original source: https://github.com/ZJONSSON/node-unzipper
 * License: MIT
 * Copyright (c) 2012 - 2013 Near Infinity Corporation
 * Commits in this fork are (c) Ziggy Jonsson (ziggy.jonsson.nyc@gmail.com)
 */

/**
 * Dates in zip file entries are stored as DosDateTime
 * Spec is here: https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-dosdatetimetofiletime
 */
export function parseDateTime(date: number, time?: number): Date {
  const day = date & 0x1f;
  const month = (date >> 5) & 0x0f;
  const year = ((date >> 9) & 0x7f) + 1980;
  const seconds = time ? (time & 0x1f) * 2 : 0;
  const minutes = time ? (time >> 5) & 0x3f : 0;
  const hours = time ? time >> 11 : 0;

  return new Date(Date.UTC(year, month - 1, day, hours, minutes, seconds));
}
