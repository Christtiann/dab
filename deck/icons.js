const React = require('react');
const { renderToStaticMarkup } = require('react-dom/server');
const Fi = require('react-icons/fi');
const sharp = require('sharp');

async function makeIcon(name, color) {
  const Comp = Fi[name];
  if (!Comp) throw new Error('icone inexistente: ' + name);
  let svg = renderToStaticMarkup(React.createElement(Comp, { color, size: 256, strokeWidth: 2 }));
  if (!svg.includes('xmlns=')) svg = svg.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');
  const buf = await sharp(Buffer.from(svg)).resize(256, 256, {
    fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 }
  }).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

async function iconSet(spec, color) {
  const out = {};
  for (const [key, name] of Object.entries(spec)) out[key] = await makeIcon(name, color);
  return out;
}

module.exports = { makeIcon, iconSet };
