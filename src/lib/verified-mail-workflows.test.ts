import { createHash } from 'crypto';
import { describe, expect, it } from 'vitest';
import {
  MAX_VERIFIED_INLINE_RESOURCES,
  VerifiedMailInputError,
  prepareReplyContent,
  providerHtmlEquivalent,
  requireNativeFolderId,
  validateContentBundle,
  verifyInlineResourceReadback,
} from './verified-mail-workflows.js';

const bytes = (value: string) => Buffer.from(value).toString('base64');
const hash = (value: string) => createHash('sha256').update(Buffer.from(value)).digest('hex');

describe('verified mail workflow primitives', () => {
  it('does not require a content bundle', () => {
    expect(validateContentBundle(undefined)).toBeUndefined();

    const result = prepareReplyContent(
      '<html><body><p>history</p></body></html>',
      'Hello',
      undefined,
      [],
      'attempt'
    );

    expect(result.resources).toEqual([]);
    expect(result.html).toContain('Hello');
    expect(result.html).toContain('<p>history</p>');
  });

  it('also permits an unchanged native reply seed when no authored content is supplied', () => {
    const seed = '<html><body><p>history</p></body></html>';
    expect(prepareReplyContent(seed, '', undefined, [], 'attempt')).toEqual({
      html: seed,
      changed: false,
      resources: [],
    });
  });

  it('validates opaque caller-provided bundles without assigning business meaning', () => {
    const one = 'image-one';
    const two = 'image-two';
    const bundle = validateContentBundle({
      bundleId: 'caller-selected-bundle',
      htmlFragment: '<img src="cid:one"><img src="cid:two">',
      inlineResources: [
        {
          contentId: 'one',
          name: 'one.png',
          contentType: 'image/png',
          contentBytesBase64: bytes(one),
          sha256: hash(one),
        },
        {
          contentId: 'two',
          name: 'two.png',
          contentType: 'image/png',
          contentBytesBase64: bytes(two),
          sha256: hash(two),
        },
      ],
    });

    expect(bundle?.inlineResources).toHaveLength(2);
    expect(MAX_VERIFIED_INLINE_RESOURCES).toBe(2);
  });

  it('preserves an opaque bundle fragment without adding a semantic wrapper', () => {
    const result = prepareReplyContent(
      '<html><body><p>history</p></body></html>',
      '',
      {
        bundleId: 'opaque',
        htmlFragment: '<p data-owner="caller">exact fragment</p>',
        inlineResources: [],
      },
      [],
      'attempt'
    );

    expect(result.html).toContain('<p data-owner="caller">exact fragment</p><p>history</p>');
    expect(result.html).not.toContain('data-mcp-content-bundle');
  });

  it('accepts only bounded Exchange HTML serialization normalization', () => {
    const expected =
      '<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">' +
      '<meta name="Generator" content="Microsoft Exchange Server">' +
      '<!-- converted from text --><style>.EmailQuote { margin-left: 1pt; }</style></head>' +
      '<body><div data-mcp-authored-content="attempt">Hello</div>' +
      '<div class="PlainText"><br>History</div></body></html>';
    const actual =
      '<html>\r\n<head>\r\n<meta content="text/html; charset=iso-8859-1" http-equiv="Content-Type">\r\n' +
      '<meta content="Microsoft Exchange Server" name="Generator">\r\n' +
      '<!-- converted from text --><style>.EmailQuote { margin-left: 1pt; }</style>\r\n</head>\r\n' +
      '<body>\r\n<div data-mcp-authored-content="attempt">Hello</div>\r\n' +
      '<div class="PlainText"><br>History</div>\r\n</body>\r\n</html>';

    expect(providerHtmlEquivalent(expected, actual)).toBe(true);
    expect(providerHtmlEquivalent(expected, actual.replace('History', 'Changed'))).toBe(false);
  });

  it('accepts only the qualified Outlook trace transition on images', () => {
    const expected = '<html><body><img src="cid:one"></body></html>';
    expect(
      providerHtmlEquivalent(
        expected,
        '<html><body><img data-outlook-trace="F:1" src="cid:one"></body></html>'
      )
    ).toBe(true);
    expect(
      providerHtmlEquivalent(
        expected,
        '<html><body><img data-outlook-trace="X:9" src="cid:one"></body></html>'
      )
    ).toBe(false);
  });

  it('fails before execution when a bundle resource hash is wrong', () => {
    expect(() =>
      validateContentBundle({
        bundleId: 'bundle',
        htmlFragment: '<img src="cid:one">',
        inlineResources: [
          {
            contentId: 'one',
            name: 'one.png',
            contentType: 'image/png',
            contentBytesBase64: bytes('one'),
            sha256: '0'.repeat(64),
          },
        ],
      })
    ).toThrowError(VerifiedMailInputError);
  });

  it('renames only current bundle content IDs when history already owns the ID', () => {
    const image = 'image';
    const result = prepareReplyContent(
      '<html><body><p><img src="cid:same">history</p></body></html>',
      '',
      {
        bundleId: 'bundle',
        htmlFragment: '<img src="cid:same">',
        inlineResources: [
          {
            contentId: 'same',
            name: 'current.png',
            contentType: 'image/png',
            contentBytesBase64: bytes(image),
            sha256: hash(image),
          },
        ],
      },
      ['same'],
      'attempt-1'
    );

    expect(result.html).toContain('cid:same-attempt-1');
    expect(result.html).toContain('cid:same">history');
    expect(result.resources[0].effectiveContentId).toBe('same-attempt-1');
  });

  it('verifies full inline bytes rather than only attachment metadata', () => {
    const image = 'image';
    const expected = {
      contentId: 'one',
      effectiveContentId: 'one',
      name: 'one.png',
      contentType: 'image/png',
      contentBytesBase64: bytes(image),
      sha256: hash(image),
    };

    expect(() =>
      verifyInlineResourceReadback(
        {
          name: expected.name,
          contentType: expected.contentType,
          isInline: true,
          contentId: expected.effectiveContentId,
          contentBytes: expected.contentBytesBase64,
        },
        expected
      )
    ).not.toThrow();

    expect(() =>
      verifyInlineResourceReadback(
        {
          name: expected.name,
          contentType: expected.contentType,
          isInline: true,
          contentId: expected.effectiveContentId,
          contentBytes: bytes('different'),
        },
        expected
      )
    ).toThrowError(VerifiedMailInputError);
  });

  it('requires an explicit native folder ID for filing', () => {
    expect(requireNativeFolderId('AAMk-folder-id')).toBe('AAMk-folder-id');
    expect(() => requireNativeFolderId('archive')).toThrowError(VerifiedMailInputError);
    expect(() => requireNativeFolderId('inbox')).toThrowError(VerifiedMailInputError);
  });
});
