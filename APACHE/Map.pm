

package Apache::Ocsinventory::Plugins::Runningdatabaseinstances;

use strict;

use Apache::Ocsinventory::Map;

$DATA_MAP{dbinstances} = {
    mask => 0,
    multi => 1,
    auto => 1,
    delOnReplace => 1,
    sortBy => 'VERSION_NAME',
    writeDiff => 0,
    cache => 0,
    fields => {
        PUBLISHER => {},
        VERSION_NAME => {},
        VERSION => {},
        EDITION => {},
        INSTANCE => {}
    }
};

1;
