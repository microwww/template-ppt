package com.github.microwww.ttp.opt;

import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class IfOperation extends Operation {

    private static final Logger logger = LoggerFactory.getLogger(IfOperation.class);

    @Override
    public void parse(ParseContext context) {
        String[] params = this.getParams();
        String join = StringUtils.join(params, " ");
        Boolean value = getValue(join, context.getDataStack(), Boolean.class);

        int start = value ? 0 : this.childrenOperations.size();
        for (int i = 0; i < this.childrenOperations.size(); i++) {
            Operation childrenOperation = this.childrenOperations.get(i);
            if (childrenOperation instanceof ElseOperation) {
                if (value) {
                    break;
                }
                start = i;
                continue;
            }
            if (i >= start) {
                childrenOperation.setParentOperations(this);
                childrenOperation.parse(context);
            }
        }
    }
}
