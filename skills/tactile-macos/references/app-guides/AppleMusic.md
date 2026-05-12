# Apple Music

## Profile

| name | workflow_mode | visual_planning | fixed_strategy |
| --- | --- | --- | --- |
| apple-music | ax-rich | false | true |

## Match Terms

- Apple Music
- Music
- 音乐
- com.apple.Music
- /System/Applications/Music.app

## Planner Guidance

- Apple Music is AX-rich enough for the default `workflow` path. Prefer `--mode auto` or explicit AX-rich execution with visual planning off unless a control is truly visual-only.
- Use the top-right/sidebar global search text field for music lookup. Do not confuse it with view-local filter controls that only narrow the current page.
- After setting the search field value, wait for the main content area to change into search results. Treat the query as unresolved until the query is still visible in the field and result sections such as `最佳结果`, `歌曲`, `专辑`, or `艺人` appear.
- Unless the user explicitly asks for the local library, prefer the `Apple Music` search scope over `你的资料库` and verify the visible scope toggle after the search loads.
- Artist matching must tolerate platform display aliases. Treat a visible artist label as a valid match when it clearly contains the user's artist name or a stable alias form of it, for example `蔡依林` matching `JOLIN蔡依林`. Do not require the platform label to be text-identical to the user's wording.
- When the initial search uses the user's plain artist name and the results surface a platform-specific display name, reuse that visible display name in one refined search if the exact song is still missing. Example: a user query with `蔡依林` may justify one retry with `JOLIN蔡依林`.
- Default decision tree:
  Search by artist when the request names an artist, for example `播放蔡依林的日不落`. Open the artist result from `最佳结果` or `艺人`, then use `歌曲排行` and `查看更多` when needed to locate the requested song from top to bottom.
  Search by song title when the request omits the artist or when a song-title search is explicitly more direct. Open the song/top result to reach the page that lists multiple versions, then choose the matching artist version. If the user did not specify an artist, play the first visible version.
- For track playback, prefer an exact title-and-artist match from `歌曲排行`, `最佳结果`, or `歌曲`. A matching album or playlist is only an intermediate navigation step, not proof that the target track is selected.
- When multiple same-title tracks exist, prefer the row whose artist matches the normalized/alias-aware target artist, even if another same-title track ranks higher in `最佳结果`.
- For artist-page navigation, do not press the artist-level `播放` button for a song-specific request. Use the artist page only as a navigation surface, then play the exact song row.
- If clicking a search result opens an album detail page, locate the exact track row in the album track list and play that row directly instead of pressing the album-level `播放` button.
- If clicking a song result opens a versions/detail page, treat that page as the source of truth for version selection and choose the row whose artist matches the request; if the request has no artist constraint, choose the first visible row.
- To start playback, double-click the track row or track container. If the first double-click only selects the row or opens the detail page, re-observe and then play the exact track row in the current view.
- Treat playback as complete only when the now-playing area changes to the expected song title and artist, and the progress display resets near `0:00`. A highlighted row alone is not enough evidence that the track is playing.
- If another song is already playing, verify that the prior now-playing metadata was replaced by the requested song before finishing.
- For queue actions such as `下一首播放` or `最后播放`, open the track `更多` menu or context menu from the exact track row, then verify the queue action label before clicking it.
- If the query is visible but results look stale or unrelated, clear the search field and set the full query again rather than appending more text to the existing search.

## Pitfalls

- The search field and filter field are not interchangeable. Using the filter field can leave the app on the current page without searching the Apple Music catalog.
- Do not reject the intended artist just because Apple Music prepends or appends a stylized stage name, romanization, or brand token around the user's name, such as `JOLIN蔡依林`.
- Do not press the artist header `播放` button when the user asked for one specific song; that can start a different popular track.
- Do not stop at the first artist search screen when the requested song is not already visible there. Open the artist page and inspect `歌曲排行` before concluding the song is missing.
- A single click on a track often only selects it. A single click on a top result can also drill into an album or artist page instead of starting playback.
- Album-level controls can start the wrong track for the user's request. Verify the specific track row before any play action when the request names a song.
- A selected track row is not enough evidence that playback switched. Always verify the now-playing title and artist.
- Apple Music can keep showing old now-playing metadata for a moment after navigation. Re-observe before concluding the requested track failed to play or already played.
