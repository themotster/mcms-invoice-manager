export function normalizeVenues(list = []) {
  return (list || []).map(item => {
    const venueId = item.venue_id != null ? Number(item.venue_id) : item.id != null ? Number(item.id) : null;
    return {
      venue_id: Number.isInteger(venueId) ? venueId : null,
      name: item.name || item.venue_name || '',
      address1: item.address1 || item.venue_address1 || '',
      address2: item.address2 || item.venue_address2 || '',
      address3: item.address3 || item.venue_address3 || '',
      town: item.town || item.venue_town || '',
      postcode: item.postcode || item.venue_postcode || '',
      is_private: Boolean(item.is_private)
    };
  });
}

export function buildVenueDraft(source = {}) {
  const venueId = source.venue_id != null ? Number(source.venue_id) : null;
  return {
    venue_id: Number.isInteger(venueId) ? venueId : null,
    name: source.name || '',
    address1: source.address1 || '',
    address2: source.address2 || '',
    address3: source.address3 || '',
    town: source.town || '',
    postcode: source.postcode || '',
    is_private: Boolean(source.is_private)
  };
}
