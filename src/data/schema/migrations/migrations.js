// This file is required for Expo/React Native SQLite migrations - https://orm.drizzle.team/quick-sqlite/expo

import journal from './meta/_journal.json';
import m0000 from './0000_misty_maverick.sql';
import m0001 from './0001_lonely_lady_deathstrike.sql';
import m0002 from './0002_cheerful_malcolm_colcord.sql';
import m0003 from './0003_zippy_bloodscream.sql';
import m0004 from './0004_aspiring_warpath.sql';
import m0005 from './0005_classy_morbius.sql';
import m0006 from './0006_lethal_frank_castle.sql';
import m0007 from './0007_sticky_agent_brand.sql';
import m0008 from './0008_lame_inertia.sql';
import m0009 from './0009_striped_donald_blake.sql';

export default {
  journal,
  migrations: {
    m0000,
    m0001,
    m0002,
    m0003,
    m0004,
    m0005,
    m0006,
    m0007,
    m0008,
    m0009,
  },
};
