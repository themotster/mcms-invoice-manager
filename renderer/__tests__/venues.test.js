import { normalizeVenues, buildVenueDraft } from '../helpers/venues';

describe('normalizeVenues', () => {
  it('converts ids to numbers and falls back to name fields', () => {
    const input = [
      {
        venue_id: '42',
        name: 'Royal Albert Hall',
        address1: 'Kensington Gore',
        town: 'London',
        postcode: 'SW7 2AP',
        is_private: 0
      },
      {
        id: '101',
        venue_name: 'Grand Ballroom',
        venue_address1: '1 Plaza Way',
        venue_town: 'Manchester',
        venue_postcode: 'M1 1AA',
        is_private: 1
      }
    ];

    const result = normalizeVenues(input);
    expect(result).toEqual([
      {
        venue_id: 42,
        name: 'Royal Albert Hall',
        address1: 'Kensington Gore',
        address2: '',
        address3: '',
        town: 'London',
        postcode: 'SW7 2AP',
        is_private: false
      },
      {
        venue_id: 101,
        name: 'Grand Ballroom',
        address1: '1 Plaza Way',
        address2: '',
        address3: '',
        town: 'Manchester',
        postcode: 'M1 1AA',
        is_private: true
      }
    ]);
  });
});

describe('buildVenueDraft', () => {
  it('builds a normalized draft with defaults', () => {
    const input = {
      venue_id: '55',
      name: 'The Roundhouse',
      address1: 'Chalk Farm Rd',
      town: 'London',
      postcode: 'NW1 8EH',
      is_private: true
    };

    expect(buildVenueDraft(input)).toEqual({
      venue_id: 55,
      name: 'The Roundhouse',
      address1: 'Chalk Farm Rd',
      address2: '',
      address3: '',
      town: 'London',
      postcode: 'NW1 8EH',
      is_private: true
    });
  });

  it('returns blank draft when no source given', () => {
    expect(buildVenueDraft()).toEqual({
      venue_id: null,
      name: '',
      address1: '',
      address2: '',
      address3: '',
      town: '',
      postcode: '',
      is_private: false
    });
  });
});
