import _ from 'lodash';

declare global {
    const LodashGS: {
        load: () => typeof _;
    };
}
